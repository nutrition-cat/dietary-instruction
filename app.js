/* ==========================================
   Web Application Logic for Nutrition Report
   ========================================== */

document.addEventListener('DOMContentLoaded', () => {
    // UI Elements
    const uploadSection = document.getElementById('upload-section');
    const reportSection = document.getElementById('report-section');
    
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const parseLoading = document.getElementById('parse-loading');
    const uploadError = document.getElementById('upload-error');
    
    const backBtn = document.getElementById('back-btn');
    const pdfBtn = document.getElementById('pdf-btn');
    
    const patientNameInput = document.getElementById('patient-name');
    const reportDateDisplay = document.getElementById('report-date-display');
    const adviceTextarea = document.getElementById('advice-textarea');

    // Global state
    let parsedData = null;
    let pfcChartInstance = null;

    // Initialize Lucide Icons
    lucide.createIcons();

    // ==========================================
    // 2. EXCEL UPLOAD & ZIP PARSING
    // ==========================================
    // Drag & Drop handlers
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.remove('dragover');
        }, false);
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length > 0) {
            handleExcelFile(files[0]);
        }
    });

    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleExcelFile(e.target.files[0]);
        }
    });

    // Excel analysis core
    async function handleExcelFile(file) {
        if (!file.name.endsWith('.xlsx')) {
            showUploadError('Excelファイル（.xlsx）をアップロードしてください。');
            return;
        }

        showLoading(true);
        showUploadError('');

        try {
            // 1. Extract images using JSZip
            const zip = await JSZip.loadAsync(file);
            const mediaFiles = {};
            
            // Extract media binaries
            for (let key in zip.files) {
                if (key.startsWith('xl/media/')) {
                    const base64 = await zip.files[key].async('base64');
                    // Key is e.g. "xl/media/image1.jpeg". Match target format in XML rels "../media/image1.jpeg"
                    const relName = key.replace('xl/media/', '../media/');
                    mediaFiles[relName] = `data:image/jpeg;base64,${base64}`;
                }
            }

            // Parse drawings to match row indices to extracted image data URLs
            const rowToImageMap = {};
            const relsFile = zip.file('xl/drawings/_rels/drawing1.xml.rels');
            const drawingFile = zip.file('xl/drawings/drawing1.xml');

            if (relsFile && drawingFile) {
                const parser = new DOMParser();
                const relsText = await relsFile.async('string');
                const relsDoc = parser.parseFromString(relsText, 'application/xml');
                const relationshipEls = relsDoc.getElementsByTagName('Relationship');
                
                const ridMap = {};
                for (let i = 0; i < relationshipEls.length; i++) {
                    const rel = relationshipEls[i];
                    ridMap[rel.getAttribute('Id')] = rel.getAttribute('Target');
                }

                const drawingText = await drawingFile.async('string');
                const drawingDoc = parser.parseFromString(drawingText, 'application/xml');
                // Support potential namespace variations by querySelectorAll on tag basename
                const anchors = drawingDoc.querySelectorAll('*|oneCellAnchor');

                anchors.forEach(anchor => {
                    const fromEl = anchor.querySelector('*|from');
                    if (fromEl) {
                        const colEl = fromEl.querySelector('*|col');
                        const rowEl = fromEl.querySelector('*|row');
                        
                        if (colEl && rowEl) {
                            const colVal = parseInt(colEl.textContent, 10);
                            const rowVal = parseInt(rowEl.textContent, 10);
                            
                            const picEl = anchor.querySelector('*|pic');
                            if (picEl) {
                                const blipEl = picEl.querySelector('*|blip');
                                if (blipEl) {
                                    // Extract embed ID (r:embed or embed)
                                    let embedRid = '';
                                    for (let attr of blipEl.attributes) {
                                        if (attr.name.endsWith('embed')) {
                                            embedRid = attr.value;
                                            break;
                                        }
                                    }
                                    
                                    const targetImage = ridMap[embedRid];
                                    const imageData = mediaFiles[targetImage];
                                    
                                    if (imageData) {
                                        // Fix 2-row layout offset:
                                        // Picture on rowVal X (0-indexed) corresponds to data record at index (rowVal - 3)
                                        // e.g. Excel Row 4 (rowVal 3) -> Data index 0
                                        const targetDataIndex = rowVal - 3;
                                        if (targetDataIndex >= 0) {
                                            rowToImageMap[targetDataIndex] = imageData;
                                        }
                                    }
                                }
                            }
                        }
                    }
                });
            }

            // 2. Parse sheet data using SheetJS
            const arrayBuffer = await readFileAsArrayBuffer(file);
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // Raw json format from worksheet
            const rawRows = XLSX.utils.sheet_to_json(worksheet);
            if (rawRows.length === 0) {
                throw new Error('Excelシート内にデータが見つかりません。');
            }

            // Extract display report date (from first row date)
            let reportDate = '2026年6月22日'; // default fallback
            if (rawRows[0]['日時']) {
                try {
                    const firstDate = new Date(rawRows[0]['日時']);
                    if (!isNaN(firstDate.getTime())) {
                        reportDate = `${firstDate.getFullYear()}年${firstDate.getMonth() + 1}月${firstDate.getDate()}日`;
                    }
                } catch(e) {}
            }

            // Process data rows and map to meals
            const meals = {
                breakfast: { items: [], kcal: 0, images: [] },
                lunch: { items: [], kcal: 0, images: [] },
                dinner: { items: [], kcal: 0, images: [] },
                other: { items: [], kcal: 0, images: [] }
            };

            const totals = {
                kcal: 0, protein: 0, fat: 0, carb: 0, fiber: 0,
                calcium: 0, iron: 0, salt: 0,
                vitA: 0, vitD: 0, vitB1: 0, vitB2: 0, vitB6: 0, vitC: 0
            };

            rawRows.forEach((row, index) => {
                // Determine meal type by time
                let mealType = 'other';
                let hours = 12; // default fallback
                
                if (row['日時']) {
                    const d = new Date(row['日時']);
                    if (!isNaN(d.getTime())) {
                        hours = d.getHours();
                    }
                }

                if (hours >= 5 && hours < 11) {
                    mealType = 'breakfast';
                } else if (hours >= 11 && hours < 16) {
                    mealType = 'lunch';
                } else if (hours >= 18 && hours < 24) {
                    mealType = 'dinner';
                } else {
                    mealType = 'other'; // e.g. 16:00-17:59, 00:00-04:59
                }

                // Map picture from rowToImageMap
                const imageSrc = rowToImageMap[index] || null;
                if (imageSrc) {
                    meals[mealType].images.push(imageSrc);
                }

                // Food details
                const foodName = row['料理名'] || '無題';
                const memo = row['メモ'] || '';
                const unit = row['単位'] || '';
                
                meals[mealType].items.push({
                    name: foodName,
                    memo: memo,
                    unit: unit
                });

                // Nutritional calculations
                const rowKcal = parseFloat(row['エネルギー（kcal）(kcal)']) || 0;
                const rowProtein = parseFloat(row['たんぱく質(g)']) || 0;
                const rowFat = parseFloat(row['脂質(g)']) || 0;
                const rowCarb = parseFloat(row['炭水化物(g)']) || 0;
                const rowFiber = parseFloat(row['食物繊維総量(g)']) || 0;
                const rowCalcium = parseFloat(row['カルシウム(mg)']) || 0;
                const rowIron = parseFloat(row['鉄(mg)']) || 0;
                const rowSalt = parseFloat(row['食塩相当量(g)']) || 0;

                const rowVitA = parseFloat(row['ビタミンA(レチノール当量)(µg)']) || 0;
                const rowVitD = parseFloat(row['ビタミンD(µg)']) || 0;
                const rowVitB1 = parseFloat(row['ビタミンB1(mg)']) || 0;
                const rowVitB2 = parseFloat(row['ビタミンB2(mg)']) || 0;
                const rowVitB6 = parseFloat(row['ビタミンB6(mg)']) || 0;
                const rowVitC = parseFloat(row['ビタミンC(mg)']) || 0;

                // Add to section totals
                meals[mealType].kcal += rowKcal;

                // Add to daily totals
                totals.kcal += rowKcal;
                totals.protein += rowProtein;
                totals.fat += rowFat;
                totals.carb += rowCarb;
                totals.fiber += rowFiber;
                totals.calcium += rowCalcium;
                totals.iron += rowIron;
                totals.salt += rowSalt;

                totals.vitA += rowVitA;
                totals.vitD += rowVitD;
                totals.vitB1 += rowVitB1;
                totals.vitB2 += rowVitB2;
                totals.vitB6 += rowVitB6;
                totals.vitC += rowVitC;
            });

            parsedData = {
                date: reportDate,
                meals: meals,
                totals: totals
            };

            // Reset upload input
            fileInput.value = '';
            showLoading(false);
            
            // Build report screen
            renderReport();
            showScreen('report-section');

        } catch (err) {
            console.error(err);
            showLoading(false);
            showUploadError(`エラーが発生しました: ${err.message || 'ファイルのパースに失敗しました。'}`);
        }
    }

    function readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = (e) => reject(new Error('ファイルの読み込みに失敗しました。'));
            reader.readAsArrayBuffer(file);
        });
    }

    // ==========================================
    // 3. REPORT RENDERING
    // ==========================================
    function renderReport() {
        if (!parsedData) return;

        // Set metadata
        reportDateDisplay.textContent = parsedData.date;
        patientNameInput.value = ''; // Reset patient name
        adviceTextarea.value = ''; // Reset advice

        const meals = parsedData.meals;
        const totals = parsedData.totals;

        // Render meal sections
        const mealTypes = ['breakfast', 'lunch', 'dinner', 'other'];
        
        mealTypes.forEach(type => {
            const container = document.getElementById(`list-${type}`);
            const imgContainer = document.getElementById(`images-${type}`);
            const data = meals[type];

            // 1. Food items
            container.innerHTML = '';
            if (data.items.length === 0) {
                container.innerHTML = `<tr><td colspan="3" class="text-muted text-center" style="font-style:italic; padding:15px;">食事データなし</td></tr>`;
            } else {
                data.items.forEach(item => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td style="font-weight: 500;">${item.name}</td>
                        <td class="text-muted">${item.memo}</td>
                        <td style="text-align: right; font-weight: 500;">${item.unit}</td>
                    `;
                    container.appendChild(tr);
                });
            }

            // 2. Images
            imgContainer.innerHTML = '';
            if (data.images.length === 0) {
                imgContainer.innerHTML = `<span class="no-image-placeholder">写真なし</span>`;
            } else {
                data.images.forEach(imgSrc => {
                    const img = document.createElement('img');
                    img.src = imgSrc;
                    img.className = 'meal-img';
                    img.alt = `${type} image`;
                    // Image click preview modal (optional, simple alert/overlay)
                    imgContainer.appendChild(img);
                });
            }

            // 3. Section kcal
            const kcalVal = document.getElementById(`val-${type}-kcal`);
            kcalVal.textContent = `${Math.round(data.kcal)} kcal`;
        });

        // Render daily totals table values
        document.getElementById('val-total-kcal').textContent = `${Math.round(totals.kcal)} kcal`;
        document.getElementById('val-total-protein').textContent = `${totals.protein.toFixed(1)} g`;
        document.getElementById('val-total-fat').textContent = `${totals.fat.toFixed(1)} g`;
        document.getElementById('val-total-carb').textContent = `${totals.carb.toFixed(1)} g`;
        document.getElementById('val-total-fiber').textContent = `${totals.fiber.toFixed(1)} g`;
        document.getElementById('val-total-calcium').textContent = `${Math.round(totals.calcium)} mg`;
        document.getElementById('val-total-iron').textContent = `${totals.iron.toFixed(1)} mg`;
        document.getElementById('val-total-salt').textContent = `${totals.salt.toFixed(1)} g`;

        // Render extra vitamins
        document.getElementById('val-total-vit-a').textContent = `${Math.round(totals.vitA)}`;
        document.getElementById('val-total-vit-d').textContent = `${totals.vitD.toFixed(1)}`;
        document.getElementById('val-total-vit-b1').textContent = `${totals.vitB1.toFixed(2)}`;
        document.getElementById('val-total-vit-b2').textContent = `${totals.vitB2.toFixed(2)}`;
        document.getElementById('val-total-vit-b6').textContent = `${totals.vitB6.toFixed(2)}`;
        document.getElementById('val-total-vit-c').textContent = `${Math.round(totals.vitC)}`;

        // Render PFC Balance
        renderPfcBalance(totals.protein, totals.fat, totals.carb);
    }

    function renderPfcBalance(protein, fat, carb) {
        // Calculate kcal values
        const pKcal = protein * 4;
        const fKcal = fat * 9;
        const cKcal = carb * 4;
        const totalPfcKcal = pKcal + fKcal + cKcal;

        let pPct = 0, fPct = 0, cPct = 0;
        
        if (totalPfcKcal > 0) {
            pPct = (pKcal / totalPfcKcal) * 100;
            fPct = (fKcal / totalPfcKcal) * 100;
            cPct = (cKcal / totalPfcKcal) * 100;
        }

        // Set legend details
        document.getElementById('pfc-p-val').textContent = `${protein.toFixed(1)}g (${Math.round(pKcal)} kcal)`;
        document.getElementById('pfc-p-pct').textContent = `${pPct.toFixed(1)}%`;
        
        document.getElementById('pfc-f-val').textContent = `${fat.toFixed(1)}g (${Math.round(fKcal)} kcal)`;
        document.getElementById('pfc-f-pct').textContent = `${fPct.toFixed(1)}%`;
        
        document.getElementById('pfc-c-val').textContent = `${carb.toFixed(1)}g (${Math.round(cKcal)} kcal)`;
        document.getElementById('pfc-c-pct').textContent = `${cPct.toFixed(1)}%`;

        // Draw Chart.js doughnut chart
        const ctx = document.getElementById('pfcChart').getContext('2d');
        
        if (pfcChartInstance) {
            pfcChartInstance.destroy();
        }

        pfcChartInstance = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['P (たんぱく質)', 'F (脂質)', 'C (炭水化物)'],
                datasets: [{
                    data: [pKcal, fKcal, cKcal],
                    backgroundColor: ['#ff6b6b', '#feca57', '#48dbfb'],
                    borderWidth: 2,
                    borderColor: '#ffffff'
                }]
            },
            options: {
                plugins: {
                    legend: {
                        display: false // We use our own customized HTML legend
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const val = context.raw;
                                const pct = ((val / totalPfcKcal) * 100).toFixed(1);
                                return `${context.label}: ${Math.round(val)} kcal (${pct}%)`;
                            }
                        }
                    }
                },
                cutout: '65%',
                responsive: true,
                maintainAspectRatio: false
            }
        });
    }

    // ==========================================
    // 4. PDF GENERATION LOGIC
    // ==========================================
    pdfBtn.addEventListener('click', () => {
        const element = document.getElementById('printable-report');
        const patientName = patientNameInput.value.trim();
        const filename = `栄養食事指導参考資料_${patientName ? patientName + '様' : '食事記録'}.pdf`;
        
        // html2pdf Options
        const opt = {
            margin: 10,
            filename: filename,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { 
                scale: 2, 
                useCORS: true,
                logging: false,
                letterRendering: true
            },
            jsPDF: { 
                unit: 'mm', 
                format: 'a4', 
                orientation: 'landscape' 
            },
            pagebreak: { mode: ['css', 'legacy'] }
        };

        // Run html2pdf
        html2pdf().set(opt).from(element).save();
    });

    backBtn.addEventListener('click', () => {
        showScreen('upload-section');
    });

    // ==========================================
    // 5. VIEW TRANSITIONS HELPERS
    // ==========================================
    function showScreen(screenId) {
        document.querySelectorAll('.screen-section').forEach(section => {
            section.classList.remove('active');
        });
        
        const targetSection = document.getElementById(screenId);
        targetSection.classList.add('active');
        
        // Relayout Lucide icons when screen changes
        lucide.createIcons();
    }

    function showLoading(isLoading) {
        if (isLoading) {
            dropZone.style.display = 'none';
            parseLoading.style.display = 'flex';
        } else {
            dropZone.style.display = 'flex';
            parseLoading.style.display = 'none';
        }
    }

    function showUploadError(message) {
        if (message) {
            uploadError.textContent = message;
            uploadError.style.display = 'flex';
        } else {
            uploadError.textContent = '';
            uploadError.style.display = 'none';
        }
    }
});
