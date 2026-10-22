"""Portable storage/UI tests; Windows message delivery is simulated explicitly."""
import importlib.util
from importlib.machinery import SourceFileLoader
import json
import os
from pathlib import Path
import queue
import tempfile
import threading
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch
import zipfile

SCRIPT = Path(__file__).with_name('hotkeys_with_template_editor.py')
spec = importlib.util.spec_from_file_location('hotkeys_editor', SCRIPT)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)


def make_docx(path, count=12):
    paragraphs = ''.join(f'<w:p><w:r><w:t>Шаблон {i}</w:t><w:br/><w:t>{{clip}}</w:t></w:r></w:p>'
                         for i in range(1, count + 1))
    xml = ('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{paragraphs}</w:body></w:document>')
    with zipfile.ZipFile(path, 'w') as archive:
        archive.writestr('word/document.xml', xml)


class StorageTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        self.store = m.TemplateStore(self.directory / 'nested' / 'templates.json')

    def test_missing_file_starts_empty(self):
        self.assertEqual(self.store.load(), {})
        self.assertFalse(self.store.path.exists())

    def test_unicode_multiline_and_placeholders_survive_restart(self):
        data = {'1': '  Здравствуйте!\n\n{clip}\t{clip1}\n  ', '9': 'Последний', '2': ''}
        self.store.save(data)
        self.assertEqual(m.TemplateStore(self.store.path).load(), data)
        self.assertEqual(json.loads(self.store.path.read_text())['version'], 1)

    def test_invalid_data_cannot_replace_saved_file(self):
        self.store.save({'1': 'Сохранено'})
        before = self.store.path.read_bytes()
        for bad in ({'10': 'Нет'}, {'1': 123}, [], {'0': ''}):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                self.store.save(bad)
            self.assertEqual(self.store.path.read_bytes(), before)

    def test_failed_atomic_replace_keeps_previous_file(self):
        self.store.save({'1': 'Старое'})
        with patch.object(m.os, 'replace', side_effect=OSError('Нет доступа')):
            with self.assertRaises(OSError):
                self.store.save({'1': 'Новое'})
        self.assertEqual(self.store.load(), {'1': 'Старое'})
        self.assertEqual(list(self.store.path.parent.glob('*.tmp')), [])

    def test_corrupt_file_is_not_silently_overwritten(self):
        self.store.path.parent.mkdir(parents=True)
        self.store.path.write_text('{ повреждённый файл', encoding='utf-8')
        before = self.store.path.read_bytes()
        with self.assertRaises(ValueError):
            self.store.load()
        self.assertEqual(self.store.path.read_bytes(), before)

    def test_import_word_preserves_order_breaks_and_limit(self):
        path = self.directory / 'old.docx'
        make_docx(path)
        imported = m.import_docx_templates(path)
        self.assertEqual(len(imported), 9)
        self.assertEqual(imported['1'], 'Шаблон 1\n{clip}')
        self.assertEqual(imported['9'], 'Шаблон 9\n{clip}')


class ClipboardTests(unittest.TestCase):
    def setUp(self):
        m._clip_history.clear()
        self.addCleanup(m._clip_history.clear)

    def test_fresh_copy_does_not_require_polling_delay(self):
        m._clip_history.appendleft('старое')
        with patch.object(m, 'pyperclip', SimpleNamespace(paste=lambda: '  новое\n')):
            result = m.substitute_clip_placeholders('{clip}|{clip1}|{clip0}|{clip4}')
        self.assertEqual(result, '  новое\n|старое|  новое\n|')

    def test_empty_current_clipboard_is_not_old_text(self):
        m._clip_history.appendleft('старое')
        with patch.object(m, 'pyperclip', SimpleNamespace(paste=lambda: '')):
            self.assertEqual(m.substitute_clip_placeholders('{clip}|{clip1}'), '|старое')


class KeyboardMappingTests(unittest.TestCase):
    def test_windows_shortcuts_ignore_latin_or_russian_layout(self):
        for code, latin, russian, action in (
            (86, 'v', 'Cyrillic_em', 'paste'),
            (67, 'c', 'Cyrillic_es', 'copy'),
            (88, 'x', 'Cyrillic_che', 'cut'),
            (65, 'a', 'Cyrillic_ef', 'select_all'),
            (90, 'z', 'Cyrillic_ya', 'undo'),
            (89, 'y', 'Cyrillic_en', 'redo'),
            (83, 's', 'Cyrillic_yeru', 'save'),
        ):
            for symbol in (latin, russian):
                with self.subTest(symbol=symbol):
                    event = SimpleNamespace(state=4, keycode=code, keysym=symbol)
                    self.assertEqual(m.editor_control_shortcut(event, windows=True), action)

    def test_altgr_does_not_trigger_editor_shortcuts(self):
        for modifier in (0x0008, 0x20000):
            event = SimpleNamespace(state=4 | modifier, keycode=86, keysym='v')
            self.assertIsNone(m.editor_control_shortcut(event, windows=True))


class WindowlessLaunchTests(unittest.TestCase):
    def test_pyw_is_a_standalone_copy_of_the_tested_program(self):
        windowless = SCRIPT.with_suffix('.pyw')
        self.assertEqual(windowless.read_bytes(), SCRIPT.read_bytes())
        loader = SourceFileLoader('hotkeys_windowless', str(windowless))
        spec = importlib.util.spec_from_loader(loader.name, loader)
        module = importlib.util.module_from_spec(spec)
        loader.exec_module(module)
        self.assertTrue(callable(module.main))
        self.assertTrue(callable(module.TemplateEditor))

    def test_console_python_relaunches_with_pythonw_and_preserves_arguments(self):
        with tempfile.TemporaryDirectory() as directory:
            python = Path(directory) / 'python.exe'
            pythonw = Path(directory) / 'pythonw.exe'
            pythonw.touch()
            with patch.object(m.sys, 'executable', str(python)), \
                    patch.object(m.sys, 'argv', [str(SCRIPT), '--example']), \
                    patch.object(m.subprocess, 'Popen') as launch:
                self.assertTrue(m.relaunch_without_console(windows=True))
            launch.assert_called_once_with(
                [str(pythonw), str(SCRIPT.resolve()), '--example'],
                creationflags=0x08000000, close_fds=True,
            )

    def test_pythonw_does_not_start_another_copy(self):
        with patch.object(m.sys, 'executable', str(Path('Python312') / 'pythonw.exe')), \
                patch.object(m.subprocess, 'Popen') as launch:
            self.assertFalse(m.relaunch_without_console(windows=True))
            launch.assert_not_called()

    def test_frozen_and_non_windows_programs_do_not_relaunch(self):
        with patch.object(m.subprocess, 'Popen') as launch:
            self.assertFalse(m.relaunch_without_console(windows=False))
            with patch.object(m.sys, 'frozen', True, create=True):
                self.assertFalse(m.relaunch_without_console(windows=True))
            launch.assert_not_called()

    def test_missing_pythonw_is_reported(self):
        with tempfile.TemporaryDirectory() as directory:
            with patch.object(m.sys, 'executable', str(Path(directory) / 'python.exe')):
                with self.assertRaisesRegex(RuntimeError, 'pythonw.exe'):
                    m.relaunch_without_console(windows=True)




class ServiceTests(unittest.TestCase):
    def setUp(self):
        m.set_templates({})
        self.addCleanup(m.set_templates, {})
        self.events = queue.Queue()
        self.service = m.HotkeyService(self.events)

    def test_only_nonempty_templates_reserve_number_keys(self):
        m.set_templates({'1': 'Текст', '2': '', '3': ' \n', '9': 'Последний'})
        bindings = self.service._bindings()
        self.assertEqual(set(bindings), {1, 9, 100, 101, 200, 201})

    def test_save_and_clear_update_registrations(self):
        api = SimpleNamespace(RegisterHotKey=Mock(return_value=1), UnregisterHotKey=Mock(return_value=1))
        with patch.object(m, 'user32', api):
            self.service._sync()
            m.set_templates({'1': 'Текст'})
            self.service._sync()
            self.assertIn(1, self.service.registered)
            m.set_templates({'1': ''})
            self.service._sync()
            self.assertNotIn(1, self.service.registered)
            api.UnregisterHotKey.assert_called_once_with(None, 1)

    def test_conflicts_are_reported_per_key(self):
        m.set_templates({'1': 'Текст', '2': 'Другой'})
        api = SimpleNamespace(RegisterHotKey=lambda hwnd, hid, mods, key: int(hid != 2))
        with patch.object(m, 'user32', api):
            self.service._sync()
        event, report = self.events.get_nowait()
        self.assertEqual(event, 'hotkeys')
        self.assertIn(1, report['registered'])
        self.assertEqual(report['failed'], {2: 'Ctrl+2'})

    def test_message_thread_registers_dispatches_and_unregisters(self):
        calls = []
        messages = iter([(m.WM_HOTKEY, 100), (m.WM_HOTKEY, 1), (m.WM_REFRESH_HOTKEYS, 0), (m.WM_QUIT, 0)])
        def get_message(pointer, *args):
            kind, key = next(messages)
            pointer._obj.message = kind
            pointer._obj.wParam = key
            return 0 if kind == m.WM_QUIT else 1
        api = SimpleNamespace(
            PeekMessageW=lambda *args: 0, GetMessageW=get_message,
            RegisterHotKey=lambda hwnd, hid, mods, key: calls.append(('register', hid, threading.get_ident())) or 1,
            UnregisterHotKey=lambda hwnd, hid: calls.append(('unregister', hid, threading.get_ident())) or 1,
        )
        m.set_templates({'1': 'Текст'})
        with patch.object(m, 'user32', api), patch.object(m, 'kernel32', SimpleNamespace(GetCurrentThreadId=threading.get_native_id)):
            self.service.start()
            self.service.thread.join(timeout=2)
        self.assertFalse(self.service.thread.is_alive())
        registered = {(hid, thread) for action, hid, thread in calls if action == 'register'}
        unregistered = {(hid, thread) for action, hid, thread in calls if action == 'unregister'}
        self.assertEqual(registered, unregistered)
        events = list(self.events.queue)
        self.assertIn(('show', None), events)
        self.assertIn(('paste', '1'), events)
        self.assertIn(('stopped', None), events)
        self.assertFalse(any(event[0] == 'error' for event in events))

    def test_hotkey_arriving_after_stop_is_not_dispatched(self):
        def get_message(pointer, *args):
            self.service.stop_requested.set()
            pointer._obj.message = m.WM_HOTKEY
            pointer._obj.wParam = 1
            return 1
        api = SimpleNamespace(
            PeekMessageW=lambda *args: 0, GetMessageW=get_message,
            RegisterHotKey=Mock(return_value=1), UnregisterHotKey=Mock(return_value=1),
        )
        m.set_templates({'1': 'Текст'})
        with patch.object(m, 'user32', api), patch.object(m, 'kernel32', SimpleNamespace(GetCurrentThreadId=threading.get_native_id)):
            self.service.start()
            self.service.thread.join(timeout=2)
        self.assertFalse(self.service.thread.is_alive())
        self.assertFalse(any(event[0] == 'paste' for event in self.events.queue))
        self.assertEqual(api.UnregisterHotKey.call_count, 5)

    def test_closed_program_does_not_start_new_workers(self):
        m._shutdown_event.set()
        self.addCleanup(m._shutdown_event.clear)
        with patch.object(m.threading, 'Thread') as worker:
            m.on_hotkey_number('1')
            m._run_excel_paste_async()
            m._run_excel_clean_async()
        worker.assert_not_called()


@unittest.skipUnless(os.environ.get('DISPLAY'), 'Tk GUI checks require a display or Xvfb')
class EditorTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        self.store = m.TemplateStore(self.directory / 'templates.json')
        m.set_templates({'1': 'Старый текст', '2': 'Второй'})
        self.addCleanup(m.set_templates, {})
        self.addCleanup(m._shutdown_event.clear)
        self.service = SimpleNamespace(refresh=Mock(), stop=Mock())
        self.events = queue.Queue()
        self.editor = m.TemplateEditor(self.store, self.service, self.events)
        self.events.put(('hotkeys', {'registered': {1, 2, 100, 101, 200, 201}, 'failed': {}}))
        self.callback_errors = []
        self.editor.report_callback_exception = lambda *args: self.callback_errors.append(args)
        self.addCleanup(self.cleanup_editor)
        self.settle()

    def cleanup_editor(self):
        try:
            exists = self.editor.winfo_exists()
        except m.tk.TclError:
            exists = False
        if exists:
            for callback in self.editor.tk.splitlist(self.editor.tk.call('after', 'info')):
                self.editor.after_cancel(callback)
            self.editor.destroy()
        self.assertEqual(self.callback_errors, [])

    def settle(self):
        self.editor.update()
        self.editor.after(110, self.editor.quit)
        self.editor.mainloop()
        self.assertEqual(self.callback_errors, [])

    def replace_text(self, text):
        self.editor.text.delete('1.0', 'end')
        self.editor.text.insert('1.0', text)
        self.settle()

    def close_window(self):
        # Execute the actual callback registered for the window's close button.
        self.editor.tk.call(self.editor.protocol('WM_DELETE_WINDOW'))

    def test_editor_is_visible_when_program_starts(self):
        self.assertEqual(self.editor.state(), 'normal')
        self.assertTrue(self.editor.winfo_viewable())
        self.assertFalse(self.editor._closing)

    def test_window_close_stops_program_even_without_reopen_hotkey(self):
        self.editor._registered.clear()
        self.close_window()
        self.assertTrue(self.editor._closing)
        self.assertTrue(m._shutdown_event.is_set())
        self.service.stop.assert_called_once()
        with self.assertRaises(m.tk.TclError):
            self.editor.winfo_exists()

    def test_cancelled_close_keeps_edits_and_hotkeys_running(self):
        self.replace_text('Несохранённый текст')
        with patch.object(m.messagebox, 'askyesnocancel', return_value=None):
            self.close_window()
        self.assertTrue(self.editor.winfo_viewable())
        self.assertFalse(self.editor._closing)
        self.assertFalse(m._shutdown_event.is_set())
        self.assertEqual(self.editor._draft['1'], 'Несохранённый текст')
        self.service.stop.assert_not_called()

    def test_close_saves_changes_before_exiting(self):
        self.replace_text('Сохранить перед выходом')
        with patch.object(m.messagebox, 'askyesnocancel', return_value=True):
            self.close_window()
        self.assertEqual(self.store.load()['1'], 'Сохранить перед выходом')
        self.assertTrue(self.editor._closing)
        self.service.stop.assert_called_once()

    def test_discard_on_close_keeps_saved_templates(self):
        self.store.save(m.get_templates())
        self.replace_text('Отбросить правки')
        with patch.object(m.messagebox, 'askyesnocancel', return_value=False):
            self.close_window()
        self.assertEqual(self.store.load()['1'], 'Старый текст')
        self.assertTrue(self.editor._closing)
        self.service.stop.assert_called_once()

    def test_failed_save_cancels_window_close(self):
        self.replace_text('Сохранение недоступно')
        with patch.object(m.messagebox, 'askyesnocancel', return_value=True), \
                patch.object(self.store, 'save', side_effect=OSError('Нет доступа')), \
                patch.object(m.messagebox, 'showerror'):
            self.close_window()
        self.assertTrue(self.editor.winfo_viewable())
        self.assertTrue(self.editor.dirty)
        self.assertFalse(m._shutdown_event.is_set())
        self.service.stop.assert_not_called()

    def test_edits_survive_switching_slots_and_only_save_publishes_them(self):
        self.replace_text('  Новый\n{clip1}\n')
        self.editor.select_template('2')
        self.settle()
        self.assertEqual(self.editor.text.get('1.0', 'end-1c'), 'Второй')
        self.editor.select_template('1')
        self.settle()
        self.assertEqual(self.editor.text.get('1.0', 'end-1c'), '  Новый\n{clip1}\n')
        self.assertEqual(m.get_templates()['1'], 'Старый текст')
        self.assertTrue(self.editor.save_templates())
        self.assertEqual(self.store.load()['1'], '  Новый\n{clip1}\n')
        self.assertEqual(m.get_templates()['1'], '  Новый\n{clip1}\n')
        self.assertFalse(self.editor.dirty)
        self.service.refresh.assert_called_once()

    def test_save_failure_keeps_draft_and_active_templates(self):
        self.replace_text('Изменение')
        with patch.object(self.store, 'save', side_effect=OSError('Нет доступа')), patch.object(m.messagebox, 'showerror') as error:
            self.assertFalse(self.editor.save_templates())
            error.assert_called_once()
        self.assertTrue(self.editor.dirty)
        self.assertEqual(m.get_templates()['1'], 'Старый текст')
        self.service.refresh.assert_not_called()



    def test_insert_placeholder_at_cursor(self):
        self.replace_text('Текст: ')
        self.editor.text.mark_set('insert', 'end-1c')
        self.editor.insert_placeholder('{clip2}')
        self.assertEqual(self.editor._draft['1'], 'Текст: {clip2}')

    def set_clipboard(self, text):
        self.editor.clipboard_clear()
        self.editor.clipboard_append(text)
        self.editor.update()

    def test_ctrl_v_pastes_once_and_updates_draft(self):
        self.replace_text('Начало: ')
        self.set_clipboard('Буфер\n{clip1}')
        self.editor.focus_force()
        self.editor.text.focus_set()
        self.editor.update()
        self.editor.text.event_generate('<Control-KeyPress-v>')
        self.settle()
        self.assertEqual(self.editor.text.get('1.0', 'end-1c'), 'Начало: Буфер\n{clip1}')
        self.assertEqual(self.editor._draft['1'], 'Начало: Буфер\n{clip1}')
        self.assertTrue(self.editor.dirty)

    def test_paste_replaces_selection_and_can_be_undone(self):
        self.replace_text('Заменить это')
        self.editor.text.edit_reset()
        self.set_clipboard('Новый\nтекст')
        self.editor.text.tag_add('sel', '1.0', 'end-1c')
        self.editor.text.event_generate('<<Paste>>')
        self.assertEqual(self.editor.text.get('1.0', 'end-1c'), 'Новый\nтекст')
        self.editor._edit_action('undo')
        self.assertEqual(self.editor.text.get('1.0', 'end-1c'), 'Заменить это')
        self.editor._edit_action('redo')
        self.assertEqual(self.editor.text.get('1.0', 'end-1c'), 'Новый\nтекст')

    def test_russian_windows_ctrl_v_and_ctrl_s_use_physical_keys(self):
        self.replace_text('Текст: ')
        self.set_clipboard('Вставка')
        self.editor._windows_keyboard = True
        paste = SimpleNamespace(state=4, keycode=86, keysym='Cyrillic_em')
        self.assertEqual(self.editor._text_control_key(paste), 'break')
        self.assertEqual(self.editor._draft['1'], 'Текст: Вставка')
        save = SimpleNamespace(state=4, keycode=83, keysym='Cyrillic_yeru')
        self.assertEqual(self.editor._text_control_key(save), 'break')
        self.assertEqual(self.store.load()['1'], 'Текст: Вставка')

    def test_empty_clipboard_does_not_delete_selection(self):
        self.replace_text('Оставить текст')
        self.editor.clipboard_clear()
        self.editor.text.tag_add('sel', '1.0', 'end-1c')
        self.editor._paste_into_editor()
        self.assertEqual(self.editor.text.get('1.0', 'end-1c'), 'Оставить текст')

    def test_edit_menu_copy_cut_and_select_all(self):
        self.replace_text('Копировать и вырезать')
        self.editor._edit_action('select_all')
        self.editor._edit_action('copy')
        self.assertEqual(self.editor.clipboard_get(), 'Копировать и вырезать')
        self.editor._edit_action('cut')
        self.assertEqual(self.editor.text.get('1.0', 'end-1c'), '')
        self.editor._edit_action('paste')
        self.assertEqual(self.editor._draft['1'], 'Копировать и вырезать')

    def test_word_import_is_a_draft_until_saved(self):
        document = self.directory / 'old.docx'
        make_docx(document, count=2)
        with patch.object(m.filedialog, 'askopenfilename', return_value=str(document)), patch.object(m.messagebox, 'askyesno', return_value=True):
            self.editor.import_from_word()
        self.assertEqual(self.editor._draft['1'], 'Шаблон 1\n{clip}')
        self.assertTrue(self.editor.dirty)
        self.assertEqual(m.get_templates()['1'], 'Старый текст')
        self.assertFalse(self.store.path.exists())
        self.assertTrue(self.editor.save_templates())
        self.assertEqual(self.store.load()['2'], 'Шаблон 2\n{clip}')

    def test_registration_conflict_and_open_events_reach_ui(self):
        self.editor.withdraw()
        self.events.put(('show', None))
        self.events.put(('hotkeys', {'registered': {1, 100, 101, 200, 201}, 'failed': {2: 'Ctrl+2'}}))
        self.settle()
        self.assertNotEqual(self.editor.state(), 'withdrawn')
        self.assertIn('Ctrl+2', self.editor.runtime_status.get())
        self.assertEqual(self.editor.template_list.item('2', 'values')[2], 'Клавиши заняты')


    def test_exit_saves_and_stops_background_service(self):
        self.replace_text('Сохранено при выходе')
        with patch.object(m.messagebox, 'askyesnocancel', return_value=True):
            self.editor.exit_program()
        self.assertTrue(self.editor._closing)
        self.assertTrue(m._shutdown_event.is_set())
        self.assertEqual(self.store.load()['1'], 'Сохранено при выходе')
        self.service.stop.assert_called_once()





if __name__ == '__main__':
    unittest.main()
