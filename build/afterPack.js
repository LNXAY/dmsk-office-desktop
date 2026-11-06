// Setzt Symbol und Programmdaten der EXE (electron-builder tut das hier nicht selbst, weil das Signieren abgeschaltet ist)
const path = require('path');
const rcedit = require('rcedit');

exports.default = async function afterPack(context) {
    const exe = path.join(context.appOutDir, context.packager.appInfo.productFilename + '.exe');
    const run = typeof rcedit === 'function' ? rcedit : rcedit.rcedit;
    await run(exe, {
        icon: path.join(context.packager.projectDir, 'build', 'icon.ico'),
        'version-string': { ProductName: 'DMSK Office', FileDescription: 'DMSK Office', CompanyName: 'DMSK Media', OriginalFilename: 'DMSK Office.exe' },
    });
};
