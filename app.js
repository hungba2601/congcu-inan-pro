let globalWorkbook = null;

document.addEventListener('DOMContentLoaded', () => {
    // alert('Ứng dụng đã sẵn sàng! Hãy chọn file rồi bấm nút Đọc File.');
    
    const fileInput = document.getElementById('fileInput');
    const classList = document.getElementById('classList');
    const selectAllBtn = document.getElementById('selectAllBtn');
    const configBtn = document.getElementById('configBtn');
    const printBtn = document.getElementById('printBtn');
    const previewWrapper = document.getElementById('previewWrapper');
    const printContainer = document.getElementById('printContainer');
    const fontSizeSetting = document.getElementById('fontSizeSetting');
    const readFileBtn = document.getElementById('readFileBtn');
    
    readFileBtn.addEventListener('click', () => {
        if (!fileInput.files || fileInput.files.length === 0) {
            alert('Bạn chưa chọn file Excel nào! Vui lòng bấm nút "Chọn tệp" trước.');
            return;
        }
        
        handleFileSelect(fileInput.files[0]);
    });

    function handleFileSelect(file) {
        try {
            classList.innerHTML = '<div class="empty-list">Đang đọc file Excel...</div>';
            
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    if (typeof XLSX === 'undefined') {
                        alert('Lỗi: Thư viện xử lý Excel chưa được tải. Vui lòng kiểm tra lại kết nối mạng (Internet) của bạn!');
                        classList.innerHTML = '<div class="empty-list" style="color: red;">Lỗi tải thư viện.</div>';
                        return;
                    }
                    
                    const data = new Uint8Array(e.target.result);
                    globalWorkbook = XLSX.read(data, {type: 'array'});
                    
                    // Populate classes
                    classList.innerHTML = '';
                    globalWorkbook.SheetNames.forEach(sheetName => {
                        const item = document.createElement('label');
                        item.className = 'class-item';
                        item.innerHTML = `
                            <input type="checkbox" value="${sheetName}" checked>
                            <span>Lớp ${sheetName}</span>
                        `;
                        classList.appendChild(item);
                    });
                } catch(err) {
                    alert('Lỗi phân tích file Excel: ' + err.message);
                    classList.innerHTML = '<div class="empty-list" style="color: red;">File Excel bị lỗi hoặc không đúng định dạng.</div>';
                }
            };
            reader.onerror = function() {
                alert('Lỗi không thể đọc dữ liệu file từ máy tính.');
                classList.innerHTML = '<div class="empty-list" style="color: red;">Không thể đọc file.</div>';
            };
            reader.readAsArrayBuffer(file);
        } catch(err) {
            alert('Lỗi khởi tạo xử lý file: ' + err.message);
        }
    }

    let allSelected = true;
    selectAllBtn.addEventListener('click', () => {
        allSelected = !allSelected;
        const checkboxes = classList.querySelectorAll('input[type="checkbox"]');
        checkboxes.forEach(cb => cb.checked = allSelected);
        selectAllBtn.textContent = allSelected ? "Bỏ chọn tất cả" : "Chọn tất cả";
    });

    configBtn.addEventListener('click', () => {
        if (!globalWorkbook) {
            alert('Vui lòng tải lên file Excel trước.');
            return;
        }

        const selectedClasses = Array.from(classList.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value);
        
        if (selectedClasses.length === 0) {
            alert('Vui lòng chọn ít nhất một lớp.');
            return;
        }

        const donVi = document.getElementById('donVi').value;
        const tenTruong = document.getElementById('tenTruong').value;
        const tieuDe = document.getElementById('tieuDe').value;
        const namHoc = document.getElementById('namHoc').value;
        const fontSize = fontSizeSetting.value + 'px';

        previewWrapper.innerHTML = '';
        printContainer.innerHTML = '';

        selectedClasses.forEach(sheetName => {
            const sheet = globalWorkbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ""});
            
            // Extract G3
            let g3Text = '';
            if (sheet['G3'] && sheet['G3'].v) {
                g3Text = sheet['G3'].v;
            } else if (data[2] && data[2][6]) {
                g3Text = data[2][6];
            }

            // Find header row index
            let headerRowIndex = 4; // default to 5th row
            for(let r = 0; r < Math.min(10, data.length); r++) {
                if(data[r] && data[r].some(cell => typeof cell === 'string' && cell.trim().toLowerCase() === 'stt')) {
                    headerRowIndex = r;
                    break;
                }
            }

            // Generate table HTML
            const tableHtml = generateTableHtml(sheet, data, headerRowIndex, fontSize);

            const pageHtml = `
                <div class="a4-paper">
                    <div class="print-header">
                        <div class="header-left">
                            <div>${donVi}</div>
                            <div class="bold">${tenTruong}</div>
                            <div>Lớp: ${sheetName}</div>
                        </div>
                        <div class="header-right">
                            <div class="bold uppercase title-lg">${tieuDe}</div>
                            <div>${namHoc}</div>
                            <div>${g3Text}</div>
                        </div>
                    </div>
                    ${tableHtml}
                </div>
            `;

            previewWrapper.insertAdjacentHTML('beforeend', pageHtml);
            printContainer.insertAdjacentHTML('beforeend', pageHtml);
        });
    });

    printBtn.addEventListener('click', () => {
        if (printContainer.innerHTML.trim() === '') {
            alert('Vui lòng Cấu Hình & Xem Trước trước khi xuất PDF.');
            return;
        }
        window.print();
    });

    fontSizeSetting.addEventListener('change', () => {
        const tables = document.querySelectorAll('.print-table');
        tables.forEach(table => {
            table.style.fontSize = fontSizeSetting.value + 'px';
        });
    });

    function generateTableHtml(sheet, data, startRowIndex, fontSize) {
        const merges = sheet['!merges'] || [];
        const colsInfo = sheet['!cols'] || [];
        const range = XLSX.utils.decode_range(sheet['!ref']);
        const maxCols = range.e.c + 1;

        const isColHidden = (c) => colsInfo[c] && colsInfo[c].hidden;

        let html = \`<table class="print-table" style="font-size: \${fontSize}">\`;
        const skipCells = {};

        for (let r = startRowIndex; r < data.length; r++) {
            // Check if row is empty
            let rowIsEmpty = true;
            for (let c = 0; c < maxCols; c++) {
                let val = (data[r] && data[r][c] !== undefined) ? data[r][c] : "";
                if (!isColHidden(c) && val !== "" && val !== null) {
                    rowIsEmpty = false;
                    break;
                }
            }
            
            // Allow empty rows if they are part of the header (sometimes they use empty rows for spacing)
            if (rowIsEmpty && r > startRowIndex + 1) continue;

            html += '<tr>';
            for (let c = 0; c < maxCols; c++) {
                if (isColHidden(c)) continue;
                if (skipCells[\`\${r},\${c}\`]) continue;

                let val = (data[r] && data[r][c] !== undefined) ? data[r][c] : "";

                let rowspan = 1;
                let colspan = 1;

                const merge = merges.find(m => m.s.r === r && m.s.c === c);
                if (merge) {
                    rowspan = merge.e.r - merge.s.r + 1;
                    colspan = 0;
                    for (let mc = merge.s.c; mc <= merge.e.c; mc++) {
                        if (!isColHidden(mc)) colspan++;
                    }

                    for (let mr = merge.s.r; mr <= merge.e.r; mr++) {
                        for (let mc = merge.s.c; mc <= merge.e.c; mc++) {
                            if (mr === r && mc === c) continue;
                            skipCells[\`\${mr},\${mc}\`] = true;
                        }
                    }
                }

                if (colspan > 0) {
                    let tag = (r === startRowIndex || r === startRowIndex + 1) ? 'th' : 'td';
                    let colspanAttr = colspan > 1 ? \` colspan="\${colspan}"\` : '';
                    let rowspanAttr = rowspan > 1 ? \` rowspan="\${rowspan}"\` : '';
                    
                    if (typeof val === 'string') {
                        val = val.replace(/\\n/g, '<br>');
                    }

                    html += \`<\${tag}\${colspanAttr}\${rowspanAttr}>\${val}</\${tag}>\`;
                }
            }
            html += '</tr>';
        }
        html += '</table>';
        return html;
    }
});
