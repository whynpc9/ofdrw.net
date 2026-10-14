#!/usr/bin/env python3
"""Mandatory original-font geometry/locator/replacement gate (currently 65 cases). Missing/invalid fixtures or any skipped case fail."""
import argparse
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import xml.etree.ElementTree as ET


def main():
    root = Path(__file__).resolve().parent.parent
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--font-output-dir', type=Path, default=root / 'artifacts/trusted-fonts')
    parser.add_argument('--results-directory', type=Path, required=True)
    args = parser.parse_args()
    results = args.results_directory.resolve()
    results.mkdir(parents=True, exist_ok=True)
    trx = results / 'trusted-fonts.trx'
    if trx.exists():
        raise ValueError('Use a fresh results directory; existing TRX must not be reused')
    loader = importlib.util.spec_from_file_location('trusted_font_prepare', root / 'scripts/prepare-trusted-fonts.py')
    module = importlib.util.module_from_spec(loader)
    loader.loader.exec_module(module)
    manifest = root / 'tests/fixtures/trusted-fonts/bundle.json'
    spec = module.read_spec(manifest)
    prepared = module.prepare(manifest, manifest.parent / spec['archive']['file'], args.font_output_dir)
    env = os.environ.copy()
    env.update(prepared['environment'])
    env.update(DOTNET_CLI_HOME=str(results / 'dotnet-home'), DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1', DOTNET_CLI_TELEMETRY_OPTOUT='1')
    Path(env['DOTNET_CLI_HOME']).mkdir(parents=True, exist_ok=True)
    command = ['dotnet', 'test', str(root / 'tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj'),
               '-c', 'Release', '--disable-build-servers', '-m:1', '/nodeReuse:false', '/p:UseSharedCompilation=false',
               '--filter', 'FullyQualifiedName~TrustedFontGeometryTests|FullyQualifiedName~KeywordLocatorRealFontTests|FullyQualifiedName~TextReplacementRealFontTests',
               '--results-directory', str(results), '--logger', 'trx;LogFileName=trusted-fonts.trx']
    (results / 'command.json').write_text(json.dumps({'command': command, 'prepared': prepared}, indent=2))
    code = subprocess.call(command, cwd=root, env=env)
    if code:
        return code
    ns = {'t': 'http://microsoft.com/schemas/VisualStudio/TeamTest/2010'}
    counters = ET.parse(trx).getroot().find('t:ResultSummary/t:Counters', ns).attrib
    if int(counters['total']) == 0 or counters['total'] != counters['passed'] or int(counters['notExecuted']) != 0:
        raise ValueError('Mandatory font gate requires actual tests and zero skips/failures')
    (results / 'summary.json').write_text(json.dumps({'counters': counters, 'bundleSha256': prepared['bundleSha256'], 'mandatory': True}, indent=2))
    return 0


if __name__ == '__main__':
    sys.exit(main())
