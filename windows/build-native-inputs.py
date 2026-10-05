"""Build the pinned Cupola ODBC fork and assemble verified Windows native inputs.
Run in an x64 Visual Studio Developer shell; no customer configuration is read.
The signed VGI extension must match the shared vgi-extensions.lock.json.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import urllib.request
import uuid
import zipfile

REPO = Path(__file__).resolve().parent.parent

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def run(*args, cwd=None, env=None):
    subprocess.run([str(arg) for arg in args], cwd=cwd, env=env, check=True)

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--vgi-extension', type=Path, required=True)
    parser.add_argument('--output', type=Path, default=REPO / 'artifacts' / 'native-inputs')
    args = parser.parse_args()
    lock = json.loads((REPO / 'windows/native-inputs.lock.json').read_text())
    vgi_lock_path = REPO / lock['vgi']['lockFile']
    vgi = json.loads(vgi_lock_path.read_text())
    if digest(args.vgi_extension) != vgi['artifacts'][lock['vgi']['platform']]['sha256']:
        raise ValueError('The VGI extension does not match the reviewed VGI extension lock.')
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    source = REPO / 'artifacts' / ('native-source-' + uuid.uuid4().hex)
    source.mkdir(parents=True)
    for key in ('haybarn', 'odbc'):
        dest = source / key
        run('git', 'clone', '--no-checkout', lock[key]['repository'], dest)
        run('git', 'checkout', '--detach', lock[key]['commit'], cwd=dest)
        actual = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=dest, text=True).strip()
        if actual != lock[key]['commit']:
            raise ValueError('Native source revision mismatch.')
    patch = REPO / lock['odbc']['patch']
    run('git', 'apply', '--check', patch, cwd=source / 'odbc')
    run('git', 'apply', patch, cwd=source / 'odbc')
    # Match the reviewed engine ABI. git describe can change as upstream adds tags
    # even when the source commit is pinned, producing an incompatible dev ABI.
    vendor_env = dict(os.environ, OVERRIDE_GIT_DESCRIBE=vgi['engineVersion'],
                      MAIN_BRANCH_VERSIONING='0', SETUPTOOLS_SCM_PRETEND_VERSION=vgi['engineVersion'],
                      SETUPTOOLS_SCM_PRETEND_HASH=lock['haybarn']['commit'][:10])
    run(sys.executable, 'vendor.py', '--duckdb', source / 'haybarn', cwd=source / 'odbc', env=vendor_env)
    version_source = (source / 'odbc/src/duckdb/src/function/table/version/pragma_version.cpp').read_text()
    for macro, value in (('DUCKDB_VERSION', vgi['engineVersion']),
                         ('DUCKDB_SOURCE_ID', lock['haybarn']['commit'][:10])):
        if f'#define {macro} "{value}"' not in version_source:
            raise ValueError('Generated Haybarn version does not match the reviewed ABI and source revision.')
    build = source / 'odbc/build/cupola'
    run('cmake', '-S', source / 'odbc', '-B', build, '-G', 'Ninja', '-DCMAKE_BUILD_TYPE=Release')
    run('cmake', '--build', build, '--target', 'haybarn_odbc', 'test_odbc', '--parallel', '4')
    run(build / 'bin/test_odbc.exe', '[cupola]')
    shutil.copy2(build / 'bin/haybarn_odbc.dll', output / 'haybarn_odbc.dll')
    wheel = source / 'haybarn-cli.whl'
    urllib.request.urlretrieve(lock['haybarn']['cliWheelUrl'], wheel)
    if digest(wheel) != lock['haybarn']['cliWheelSha256']:
        raise ValueError('Haybarn wheel checksum mismatch.')
    with zipfile.ZipFile(wheel) as archive:
        matches = [name for name in archive.namelist() if name.endswith('/_bin/haybarn.exe')]
        if len(matches) != 1:
            raise ValueError('The Haybarn wheel must contain exactly one native CLI.')
        (output / 'haybarn.exe').write_bytes(archive.read(matches[0]))
    if digest(output / 'haybarn.exe') != lock['haybarn']['cliSha256']:
        raise ValueError('Haybarn CLI checksum mismatch.')
    shutil.copy2(args.vgi_extension, output / 'vgi.duckdb_extension')
    (output / 'provenance.json').write_text(json.dumps({
        'lock': lock,
        'lockSha256': digest(REPO / 'windows/native-inputs.lock.json'),
        'patchSha256': digest(patch),
        'vgiLockSha256': digest(vgi_lock_path),
        'files': {name: digest(output / name) for name in ('haybarn.exe', 'haybarn_odbc.dll', 'vgi.duckdb_extension')},
    }, indent=2) + '\n')
    print('Verified native inputs are ready:', output)

if __name__ == '__main__':
    main()
