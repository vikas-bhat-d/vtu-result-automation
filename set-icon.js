const resedit = require('resedit');
const fs = require('fs');

const exePath  = 'dist/vtu-result.exe';
const icoPath  = 'public/vtu-icon.ico';

const data = fs.readFileSync(exePath);
const ex   = resedit.NtExecutable.from(data);
const res  = resedit.NtExecutableResource.from(ex);
const ico  = resedit.Data.IconFile.from(fs.readFileSync(icoPath));

const groups = resedit.Resource.IconGroupEntry.fromEntries(res.entries);
resedit.Resource.IconGroupEntry.replaceIconsForResource(
  res.entries,
  groups[0].id,
  1033,
  ico.icons.map(i => i.data)
);

res.outputResource(ex);
fs.writeFileSync(exePath, Buffer.from(ex.generate()));
console.log('Icon set successfully on', exePath);
