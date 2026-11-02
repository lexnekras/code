"""Publish build products/logs using Git; no GitHub REST API is needed by the agent."""
from __future__ import annotations
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import tempfile

ROOT = Path(__file__).resolve().parents[2]
OUTPUT = ROOT / 'windows-output'
REF = 'refs/heads/codex/hotkeys-windows-exe'


def git(*arguments, env=None):
    return subprocess.check_output(['git', '-C', str(ROOT), *arguments], text=True, env=env).strip()


def main():
    OUTPUT.mkdir(exist_ok=True)
    phase = os.environ['BUILD_PHASE']
    report_path = OUTPUT / 'build-report.json'
    report = json.loads(report_path.read_text(encoding='utf-8')) if report_path.exists() else {}
    success = phase == 'complete' and os.environ.get('BUILD_JOB_STATUS') == 'success' and report.get('status') == 'success'
    status = {
        'status': 'running' if phase == 'running' else ('success' if success else 'failed'),
        'build_commit': os.environ['GITHUB_SHA'],
        'run_url': f"https://github.com/{os.environ['GITHUB_REPOSITORY']}/actions/runs/{os.environ['GITHUB_RUN_ID']}",
        'updated_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(),
    }
    if success:
        status['exe_sha256'] = hashlib.sha256((OUTPUT / 'Hotkeys.exe').read_bytes()).hexdigest()
    (OUTPUT / 'build-status.json').write_text(json.dumps(status, indent=2), encoding='utf-8')
    dependencies = ROOT / 'dependency-install.log'
    if dependencies.exists():
        (OUTPUT / 'dependency-install.log').write_bytes(dependencies.read_bytes())
    remote = git('ls-remote', '--heads', 'origin', REF)
    parent = remote.split()[0] if remote else None
    if parent:
        git('fetch', '--no-tags', '--depth=1', 'origin', REF)
    with tempfile.TemporaryDirectory(prefix='hotkeys-publish-') as temporary:
        environment = os.environ.copy()
        environment['GIT_INDEX_FILE'] = str(Path(temporary) / 'index')
        git('read-tree', '--empty', env=environment)
        for file in sorted(OUTPUT.iterdir()):
            if not file.is_file() or (not success and file.suffix.lower() in {'.exe', '.zip'}):
                continue
            blob = git('hash-object', '-w', str(file))
            git('update-index', '--add', '--cacheinfo', f'100644,{blob},{file.name}', env=environment)
        tree = git('write-tree', env=environment)
        command = ['-c', 'user.name=github-actions[bot]', '-c',
                   'user.email=41898282+github-actions[bot]@users.noreply.github.com',
                   'commit-tree', tree, '-m', f"Windows build {status['status']} for {os.environ['GITHUB_SHA'][:12]}"]
        if parent:
            command.extend(['-p', parent])
        commit = git(*command, env=environment)
        git('update-ref', REF, commit)
        git('push', 'origin', f'{REF}:{REF}')
    print(json.dumps(status), flush=True)


if __name__ == '__main__':
    main()
