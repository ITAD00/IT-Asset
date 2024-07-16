document.addEventListener('DOMContentLoaded', function() {
    const assetForm = document.getElementById('assetForm');
    const assetTableBody = document.querySelector('#assetTable tbody');
    const exportPdfButton = document.getElementById('exportPdf');
    const exportExcelButton = document.getElementById('exportExcel');
    const prevPageButton = document.getElementById('prevPage');
    const nextPageButton = document.getElementById('nextPage');

    const assets = [];
    const assetsPerPage = 5;
    let currentPage = 1;

    assetForm.addEventListener('submit', function(event) {
        event.preventDefault();

        const assetHolderName = document.getElementById('assetHolderName').value;
        const assetType = document.getElementById('assetType').value;
        const assetStatus = document.getElementById('assetStatus').value;
        const hostname = document.getElementById('hostname').value;
        const productBrand = document.getElementById('productBrand').value;
        const modelNumber = document.getElementById('modelNumber').value;
        const serialNumber = document.getElementById('serialNumber').value;
        const purchaseDate = document.getElementById('purchaseDate').value;
        const os = document.getElementById('os').value;
        const ram = document.getElementById('ram').value;
        const storage = document.getElementById('storage').value;
        const processor = document.getElementById('processor').value;
        const warranty = document.getElementById('warranty').value;
        const vendor = document.getElementById('vendor').value;

        const asset = {
            id: assets.length + 1,
            holderName: assetHolderName,
            type: assetType,
            status: assetStatus,
            hostname: hostname,
            productBrand: productBrand,
            modelNumber: modelNumber,
            serialNumber: serialNumber,
            purchaseDate: purchaseDate,
            configuration: `OS: ${os}, RAM: ${ram}, Storage: ${storage}, Processor: ${processor}`,
            warranty: warranty,
            vendor: vendor,
            age: calculateAssetAge(purchaseDate)
        };

        assets.push(asset);
        updateAssetTable();
        assetForm.reset();
    });

    function updateAssetTable() {
        assetTableBody.innerHTML = '';

        const startIndex = (currentPage - 1) * assetsPerPage;
        const endIndex = startIndex + assetsPerPage;
        const paginatedAssets = assets.slice(startIndex, endIndex);

        paginatedAssets.forEach(asset => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${asset.id}</td>
                <td>${asset.holderName}</td>
                <td>${asset.type}</td>
                <td>${asset.status}</td>
                <td>${asset.hostname}</td>
                <td>${asset.productBrand}</td>
                <td>${asset.modelNumber}</td>
                <td>${asset.serialNumber}</td>
                <td>${asset.purchaseDate}</td>
                <td>${asset.configuration}</td>
                <td>${asset.warranty}</td>
                <td>${asset.vendor}</td>
            `;
            assetTableBody.appendChild(row);
        });
    }

    function calculateAssetAge(purchaseDate) {
        const purchase = new Date(purchaseDate);
        const now = new Date();
        const age = now.getFullYear() - purchase.getFullYear();
        return age + (now.getMonth() >= purchase.getMonth() && now.getDate() >= purchase.getDate() ? '' : ' - 1') + ' year(s)';
    }

    prevPageButton.addEventListener('click', function() {
        if (currentPage > 1) {
            currentPage--;
            updateAssetTable();
        }
    });

    nextPageButton.addEventListener('click', function() {
        if (currentPage * assetsPerPage < assets.length) {
            currentPage++;
            updateAssetTable();
        }
    });

    exportPdfButton.addEventListener('click', function() {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();

        doc.text("Assets List", 14, 16);
        doc.autoTable({
            startY: 20,
            head: [['ID', 'Holder Name', 'Type', 'Status', 'Hostname', 'Product Brand', 'Model Number', 'Serial Number', 'Purchase Date', 'Configuration', 'Warranty', 'Vendor']],
            body: assets.map(asset => [
                asset.id,
                asset.holderName,
                asset.type,
                asset.status,
                asset.hostname,
                asset.productBrand,
                asset.modelNumber,
                asset.serialNumber,
                asset.purchaseDate,
                asset.configuration,
                asset.warranty,
                asset.vendor
            ]),
            theme: 'grid',
            styles: {
                fillColor: [173, 216, 230],
                textColor: [0, 0, 0],
                fontSize: 10
            }
        });

        doc.save('assets.pdf');
    });

    exportExcelButton.addEventListener('click', function() {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(assets.map(asset => ({
            ID: asset.id,
            HolderName: asset.holderName,
            Type: asset.type,
            Status: asset.status,
            Hostname: asset.hostname,
            ProductBrand: asset.productBrand,
            ModelNumber: asset.modelNumber,
            SerialNumber: asset.serialNumber,
            PurchaseDate: asset.purchaseDate,
            Configuration: asset.configuration,
            Warranty: asset.warranty,
            Vendor: asset.vendor
        })));

        // Apply the "Table Style Medium 6" to the worksheet
        const table = XLSX.utils.sheet_to_json(ws, { header: 1 });
        XLSX.utils.sheet_add_aoa(ws, [["ID", "Holder Name", "Type", "Status", "Hostname", "Product Brand", "Model Number", "Serial Number", "Purchase Date", "Configuration", "Warranty", "Vendor"]], { origin: "A1" });
        XLSX.utils.book_append_sheet(wb, ws, 'Assets');

        ws['!ref'] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 11, r: assets.length } });
        const tableRef = XLSX.utils.decode_range(ws['!ref']);
        tableRef.s.r = 1; // Skip header row
        ws['!autofilter'] = { ref: XLSX.utils.encode_range(tableRef) };

        const format = {
            border: {
                style: 'thin',
                color: { rgb: "CCCCCC" }
            },
            fill: {
                fgColor: { rgb: "B6DDE8" }
            }
        };
        for (let R = 1; R <= assets.length; ++R) {
            for (let C = 0; C <= 11; ++C) {
                const cell_address = { c: C, r: R };
                const cell_ref = XLSX.utils.encode_cell(cell_address);
                if (!ws[cell_ref]) continue;
                ws[cell_ref].s = format;
            }
        }

        XLSX.writeFile(wb, 'assets.xlsx');
    });
});
