#!/usr/bin/env python3
"""Validate local consumer provenance, default CLI isolation and every rendered sample page."""
import base64,hashlib,json,pathlib,subprocess,sys,xml.etree.ElementTree as E,zipfile

root=pathlib.Path(sys.argv[1]); version=sys.argv[2]; files=root/'files'
assets=json.loads(pathlib.Path(sys.argv[3]).read_text())
assert not any(v['type']=='project' for v in assets['libraries'].values()), 'Consumer must use packages only'
assert f'Ofdrw.Net.Signatures.SesInterop/{version}' in assets['libraries']
cli_deps=pathlib.Path(sys.argv[4]); cli=json.loads(cli_deps.read_text())
assert not any('SesInterop' in k or 'BouncyCastle' in k for k in cli['libraries']), 'CLI must not load optional crypto'
data=json.loads((files/'functional.json').read_text())
assert data['capabilities']['coreBuiltInSesSm2'] is False
assert len(data['samples'])==4
for sample in data['samples']:
 assert sample['referenceIntegrityValid'] and sample['explicitTestFullyValid'] and not sample['defaultFullyValid'] and not sample['wrongPinFullyValid']
 assert sample['defaultCryptographicStatus']=='Unsupported' and sample['textEqual'] and sample['originalEntriesPreserved']
 text=(files/(sample['name']+'-cli.txt')).read_text()
 assert 'signed-value Unsupported' in text and 'reference integrity valid' in text and 'fully valid' not in text

packages={}
for package in sorted((root/'packages').rglob('*.nupkg')):
 with zipfile.ZipFile(package) as z:
  spec=E.fromstring(z.read(next(n for n in z.namelist() if n.endswith('.nuspec'))))
  values={n.tag.rsplit('}',1)[-1]:n.text for n in spec.iter()}
  assert values['version']==version
  identity=values['id']+'/'+version
  content_hash=base64.b64encode(hashlib.sha512(package.read_bytes()).digest()).decode()
  if identity in assets['libraries']: assert assets['libraries'][identity]['sha512']==content_hash, 'Consumed package hash mismatch'
  if identity in cli['libraries'] and cli['libraries'][identity].get('sha512'):
   assert cli['libraries'][identity]['sha512']=='sha512-'+content_hash, 'CLI package hash mismatch'
  if values['id']=='Ofdrw.Net.Signatures.SesInterop':
   assert 'licenses/BouncyCastle-LICENSE.md' in z.namelist()
   deps=[n for n in spec.iter() if n.tag.rsplit('}',1)[-1]=='dependency' and n.get('id')=='BouncyCastle.Cryptography']
   assert deps and all(n.get('version')=='[2.6.2]' for n in deps)
   assert all(f'lib/{tfm}/Ofdrw.Net.Signatures.SesInterop.dll' in z.namelist() for tfm in ('netstandard2.0','netstandard2.1'))
  elif values['id'] in ('Ofdrw.Net.Core','Ofdrw.Net.Packaging','Ofdrw.Net.Signatures','Ofdrw.Net.Converter','Ofdrw.Net.Cli'):
   assert not any(n.get('id','') in ('Ofdrw.Net.Signatures.SesInterop','BouncyCastle.Cryptography') for n in spec.iter())
 packages[str(package.relative_to(root))]={'bytes':package.stat().st_size,'sha256':hashlib.sha256(package.read_bytes()).hexdigest()}
assert len(packages)==12
cli_package=root/'packages/sdk'/f'Ofdrw.Net.Cli.{version}.nupkg'
installed_cli_dlls={}
with zipfile.ZipFile(cli_package) as z:
 prefix='tools/net10.0/any/'
 assert z.read(prefix+'Ofdrw.Net.Cli.deps.json')==cli_deps.read_bytes()
 for name in z.namelist():
  if name.startswith(prefix+'Ofdrw.Net.') and name.endswith('.dll'):
   installed=cli_deps.parent/pathlib.PurePosixPath(name).name
   assert installed.read_bytes()==z.read(name), 'Installed CLI DLL differs from local package'
   installed_cli_dlls[installed.name]=hashlib.sha256(installed.read_bytes()).hexdigest()
 for package in (root/'packages/sdk').glob('*.nupkg'):
  if package==cli_package: continue
  with zipfile.ZipFile(package) as sdk:
   names=[n for n in sdk.namelist() if n.startswith('lib/netstandard2.1/') and n.endswith('.dll')]
   for name in names:
    dll=pathlib.PurePosixPath(name).name
    assert sdk.read(name)==z.read(prefix+dll), 'CLI SDK DLL differs from corresponding SDK package'
(files/'installed-cli-dlls.json').write_text(json.dumps(installed_cli_dlls,indent=2)+'\n')
(files/'package-manifest.json').write_text(json.dumps({'version':version,'packages':packages,'consumerPackageOnly':True,'cliOptionalDependencyAbsent':True},indent=2)+'\n')
(files/'consumer-assets.json').write_text(json.dumps(assets,indent=2)+'\n')
(files/'cli-deps.json').write_text(json.dumps(cli,indent=2)+'\n')

pages={}
for pdf in sorted(files.glob('*.pdf')):
 directory=files/'pages'/pdf.stem; directory.mkdir(parents=True,exist_ok=True)
 subprocess.run(['pdftoppm','-r','96','-png',str(pdf),str(directory/'page')],check=True)
 images=sorted(directory.glob('page-*.png'))
 pages[pdf.stem]=[{'file':str(p.relative_to(files)),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in images]
 assert images
for sample in data['samples']:
 actual=pages[sample['name']]; baseline=pages['baseline-'+sample['mode']]
 assert len(actual)==sample['pages']==len(baseline)
 assert [p['sha256'] for p in actual]==[p['sha256'] for p in baseline], 'Signing changed rendered page pixels'
(files/'rendered-pages.json').write_text(json.dumps({'dpi':96,'pages':pages,'signedPixelsEqualBaseline':True,'previewAccepted':False},indent=2)+'\n')
print('Local 12-package provenance, default CLI exit 2, four test signatures and all page PNG comparisons passed; Preview remains separate.')
