const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const JSZip = require('jszip');

const packageRoot = path.resolve(__dirname, '..');
const fixtureDir = path.join(__dirname, 'fixtures');

function loadLuckyFileFromDist() {
    const bundlePath = path.join(packageRoot, 'dist', 'luckyexcel.cjs.js');
    const bundle = `${fs.readFileSync(bundlePath, 'utf8')}\nmodule.exports={LuckyExcel:module.exports,LuckyFile,ReadXml,HandleZip};`;

    function DummyDOMParser() {}
    DummyDOMParser.prototype.parseFromString = function () {
        return { children: [] };
    };

    function DummyXMLSerializer() {}
    DummyXMLSerializer.prototype.serializeToString = function () {
        return '';
    };

    function DummyFileReader() {}

    const sandbox = {
        module: { exports: {} },
        exports: {},
        require,
        console,
        navigator: { userAgent: 'node' },
        document: {
            createElement() {
                return {
                    style: {},
                    click() {},
                    addEventListener() {},
                    remove() {},
                };
            },
            body: {
                appendChild() {},
            },
        },
        URL: {
            createObjectURL() {
                return 'blob:fixture';
            },
            revokeObjectURL() {},
        },
        Blob,
        setTimeout,
        DOMParser: DummyDOMParser,
        XMLSerializer: DummyXMLSerializer,
        FileReader: DummyFileReader,
    };

    vm.runInNewContext(bundle, sandbox, { filename: bundlePath });
    return sandbox.module.exports.LuckyFile;
}

async function unzipWorkbookFixture(fileName) {
    const filePath = path.join(fixtureDir, fileName);
    const workbookBuffer = fs.readFileSync(filePath);
    const zip = await JSZip.loadAsync(workbookBuffer);
    const entries = {};

    for (const [entryName, file] of Object.entries(zip.files)) {
        if (file.dir) {
            continue;
        }

        const extension = path.extname(entryName).slice(1).toLowerCase();
        let dataType = 'string';

        if (['png', 'jpeg', 'jpg', 'gif', 'bmp', 'tif', 'webp'].includes(extension)) {
            dataType = 'base64';
        } else if (extension === 'emf') {
            dataType = 'arraybuffer';
        }

        let contents = await file.async(dataType);
        if (dataType === 'base64') {
            contents = `data:image/${extension};base64,${contents}`;
        }

        entries[entryName] = contents;
    }

    return entries;
}

async function parseFixture(fileName) {
    const LuckyFile = loadLuckyFileFromDist();
    const entries = await unzipWorkbookFixture(fileName);
    return new LuckyFile(entries, fileName).ParseObject();
}

test('imports appointment-schedule.xlsx without throwing on theme-based conditional formatting colors', async () => {
    const workbook = await parseFixture('appointment-schedule.xlsx');

    assert.ok(workbook);
    assert.ok(Array.isArray(workbook.sheets));
    assert.ok(workbook.sheets.length > 0);
    assert.ok(workbook.sheets.some((sheet) => sheet.name === 'Example'));
    assert.ok(
        workbook.sheets.some(
            (sheet) => Array.isArray(sheet.conditionalFormatting) && sheet.conditionalFormatting.length > 0
        )
    );
});

test('still imports known-good control workbook after theme color normalization changes', async () => {
    const workbook = await parseFixture('biweekly-work-schedule.xlsx');

    assert.ok(workbook);
    assert.ok(Array.isArray(workbook.sheets));
    assert.ok(workbook.sheets.some((sheet) => sheet.name === 'Week 1-2'));
});
