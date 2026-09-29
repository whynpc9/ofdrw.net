#!/usr/bin/env python3
"""Fail when a new public SDK member lacks XML docs; baseline existing debt by symbol."""
import argparse
import json
import os
from pathlib import Path
import re
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
BASELINE = ROOT / 'docs' / 'public-api-cs1591-baseline.json'
DIAGNOSTIC = re.compile(r'^(.*?\.cs)\(\d+,\d+\): warning CS1591: Missing XML comment for publicly visible (?:type or member) [\'\"](.+?)[\'\"]')


def parse(output):
    result = set()
    for line in output.splitlines():
        match = DIAGNOSTIC.search(line)
        if not match:
            continue
        path = Path(match.group(1))
        try:
            relative = path.resolve().relative_to(ROOT).as_posix()
        except ValueError:
            continue
        if relative.startswith('src/'):
            result.add(f'{relative}|{match.group(2)}')
    return sorted(result)


def unparsed_diagnostics(output):
    return [line for line in output.splitlines()
            if re.search(r'(?:warning|error) CS1591:', line) and not DIAGNOSTIC.search(line)]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--update-baseline', action='store_true', help='explicitly approve current historical debt')
    parser.add_argument('--package-version', help='keep the release build stamped with the tag version')
    args = parser.parse_args()
    env = os.environ.copy()
    env.setdefault('DOTNET_CLI_HOME', str(ROOT / 'artifacts' / 'dotnet-home'))
    env['DOTNET_SKIP_FIRST_TIME_EXPERIENCE'] = '1'
    env['DOTNET_CLI_TELEMETRY_OPTOUT'] = '1'
    command = ['dotnet', 'build', str(ROOT / 'Ofdrw.Net.sln'), '-c', 'Release', '--no-restore',
               '-t:Rebuild', '-p:EnforcePublicApiDocs=true', '--disable-build-servers',
               '-m:1', '/nodeReuse:false', '/p:UseSharedCompilation=false', '--nologo']
    if args.package_version:
        command.extend([f'-p:Version={args.package_version}', f'-p:PackageVersion={args.package_version}'])
    completed = subprocess.run(command, cwd=ROOT, env=env, text=True, stdout=subprocess.PIPE,
                               stderr=subprocess.STDOUT, check=False)
    actual = parse(completed.stdout)
    if completed.returncode:
        print(completed.stdout, file=sys.stderr)
        raise ValueError(f'Documentation build failed with exit code {completed.returncode}.')
    unparsed = unparsed_diagnostics(completed.stdout)
    if unparsed:
        raise ValueError('Unrecognized CS1591 output format: ' + '\n'.join(unparsed[:5]))
    if args.update_baseline:
        BASELINE.write_text(json.dumps(actual, ensure_ascii=False, indent=2) + '\n')
        print(f'Wrote {len(actual)} historical CS1591 symbols to {BASELINE}.')
        return
    baseline = set(json.loads(BASELINE.read_text()))
    if baseline and not actual:
        raise ValueError('No CS1591 diagnostics were captured despite a non-empty historical baseline.')
    added = sorted(set(actual) - baseline)
    removed = sorted(baseline - set(actual))
    if added:
        print('New undocumented public API:\n' + '\n'.join(added), file=sys.stderr)
        raise ValueError(f'{len(added)} new CS1591 diagnostic(s).')
    if removed:
        print('Documented historical members:\n' + '\n'.join(removed))
    print(f'Public API documentation gate passed: {len(actual)} historical symbols, {len(removed)} cleared.')


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, json.JSONDecodeError) as error:
        print(f'Public API documentation gate failed: {error}', file=sys.stderr)
        sys.exit(1)
