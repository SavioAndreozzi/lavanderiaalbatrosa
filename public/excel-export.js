// Update the supplied workbook directly to preserve its layout and both sheets.
async function createLaundryWorkbook(records, pezzoItems, pesoItems) {
    const response = await fetch('/assets/GRIGLIATO.xlsx');
    if (!response.ok) throw new Error('Modello Excel non disponibile');
    const zip = await JSZip.loadAsync(await response.arrayBuffer());
    const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const parse = async path => new DOMParser().parseFromString(await zip.file(path).async('string'), 'application/xml');
    const grid = await parse('xl/worksheets/sheet1.xml');
    const matrix = await parse('xl/worksheets/sheet2.xml');
    const strings = Array.from((await parse('xl/sharedStrings.xml')).getElementsByTagNameNS(ns, 'si'), el => el.textContent);
    const cells = doc => new Map(Array.from(doc.getElementsByTagNameNS(ns, 'c'), cell => [cell.getAttribute('r'), cell]));
    const gridCells = cells(grid);
    const read = cell => {
        if (!cell) return '';
        const value = cell.getElementsByTagNameNS(ns, 'v')[0]?.textContent || '';
        return cell.getAttribute('t') === 's' ? strings[Number(value)] : value;
    };
    const column = number => {
        let name = '';
        for (; number; number = Math.floor((number - 1) / 26)) name = String.fromCharCode(65 + (number - 1) % 26) + name;
        return name;
    };
    const roster = new Map();
    for (let row = 7; row <= 159; row++) {
        const cod = read(gridCells.get(`A${row}`));
        if (cod) roster.set(cod.trim().toUpperCase(), { cod, cognome: read(gridCells.get(`B${row}`)), nome: read(gridCells.get(`C${row}`)) });
    }
    for (const record of records) roster.set(record.cod.trim().toUpperCase(), record);
    const sorted = [...roster.values()].sort((a, b) => a.cod.trim().localeCompare(b.cod.trim(), 'it', { numeric: true, sensitivity: 'base' }));
    const sheetData = grid.getElementsByTagNameNS(ns, 'sheetData')[0];
    const sampleRow = grid.getElementsByTagNameNS(ns, 'row')[6];
    const lastRow = Math.max(159, sorted.length + 6);
    const rows = new Map(Array.from(sheetData.children, row => [Number(row.getAttribute('r')), row]));
    function getCell(rowNumber, colNumber) {
        const address = `${column(colNumber)}${rowNumber}`;
        if (gridCells.has(address)) return gridCells.get(address);
        let row = rows.get(rowNumber);
        if (!row) {
            row = grid.createElementNS(ns, 'row');
            row.setAttribute('r', rowNumber);
            if (sampleRow.hasAttribute('ht')) row.setAttribute('ht', sampleRow.getAttribute('ht'));
            const next = [...sheetData.children].find(el => Number(el.getAttribute('r')) > rowNumber);
            sheetData.insertBefore(row, next || null);
            rows.set(rowNumber, row);
        }
        const cell = grid.createElementNS(ns, 'c');
        cell.setAttribute('r', address);
        const sample = gridCells.get(`${column(colNumber)}7`);
        if (sample?.hasAttribute('s')) cell.setAttribute('s', sample.getAttribute('s'));
        const next = [...row.children].find(el => {
            const letters = el.getAttribute('r').replace(/\d/g, '');
            return letters.length > column(colNumber).length || (letters.length === column(colNumber).length && letters > column(colNumber));
        });
        row.insertBefore(cell, next || null);
        gridCells.set(address, cell);
        return cell;
    }
    function write(cell, value, formula) {
        cell.replaceChildren();
        cell.removeAttribute('t');
        if (formula) {
            const f = cell.ownerDocument.createElementNS(ns, 'f');
            f.textContent = formula;
            cell.append(f);
        }
        if (typeof value === 'string' && !formula) {
            cell.setAttribute('t', 'inlineStr');
            const inline = cell.ownerDocument.createElementNS(ns, 'is');
            const text = cell.ownerDocument.createElementNS(ns, 't');
            text.textContent = value;
            inline.append(text);
            cell.append(inline);
        } else if (value !== null) {
            const v = cell.ownerDocument.createElementNS(ns, 'v');
            v.textContent = value;
            cell.append(v);
        }
    }
    const totals = Array(47).fill(0);
    for (let row = 7; row <= lastRow; row++) {
        const record = sorted[row - 7];
        const values = record ? [record.cod, record.cognome, record.nome,
            ...pezzoItems.map(item => Number(record.pezzo?.[item.id]) || null), null,
            ...pesoItems.map(item => Number(record.peso?.[item.id]) || null), null] : [];
        for (let col = 1; col <= 47; col++) {
            const value = values[col - 1] ?? null;
            write(getCell(row, col), value);
            if (typeof value === 'number') totals[col - 1] += value;
        }
    }
    for (let col = 4; col <= 47; col++) {
        // Bag counts are not collected by the app, so leave those fields empty.
        write(getCell(6, col), col === 23 || col === 47 ? null : totals[col - 1],
            `SUM(${column(col)}7:${column(col)}${lastRow})`);
    }
    for (const cell of cells(matrix).values()) {
        const formula = cell.getElementsByTagNameNS(ns, 'f')[0]?.textContent;
        const match = formula?.match(/^Grigliato!([A-Z]+)([56])$/);
        if (match) {
            const source = gridCells.get(`${match[1]}${match[2]}`);
            if (match[2] === '5') continue; // Keep the original cached garment labels.
            const value = read(source);
            write(cell, value === '' ? null : Number(value), formula);
        }
    }
    // Replace the sample service date with the date of this export.
    const today = new Date();
    const excelDate = Math.round((Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()) - Date.UTC(1899, 11, 30)) / 86400000);
    write(cells(matrix).get('B9'), excelDate);
    const workbook = await parse('xl/workbook.xml');
    const calc = workbook.getElementsByTagNameNS(ns, 'calcPr')[0];
    calc.setAttribute('fullCalcOnLoad', '1');
    calc.setAttribute('forceFullCalc', '1');
    const dimension = grid.getElementsByTagNameNS(ns, 'dimension')[0];
    if (lastRow > 204) dimension.setAttribute('ref', `A1:ML${lastRow}`);
    const serializer = new XMLSerializer();
    zip.file('xl/worksheets/sheet1.xml', serializer.serializeToString(grid));
    zip.file('xl/worksheets/sheet2.xml', serializer.serializeToString(matrix));
    zip.file('xl/workbook.xml', serializer.serializeToString(workbook));
    // An old calculation chain may still refer to shared formulas replaced above.
    zip.remove('xl/calcChain.xml');
    for (const path of ['[Content_Types].xml', 'xl/_rels/workbook.xml.rels']) {
        const doc = await parse(path);
        for (const node of [...doc.documentElement.children]) {
            if ((node.getAttribute('PartName') || node.getAttribute('Target') || '').includes('calcChain')) node.remove();
        }
        zip.file(path, serializer.serializeToString(doc));
    }
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', compression: 'DEFLATE' });
}
