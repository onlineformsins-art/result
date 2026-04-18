document.addEventListener('DOMContentLoaded', () => {
    // Set current year in footer
    document.getElementById('year').textContent = new Date().getFullYear();

    const searchBtn = document.getElementById('searchBtn');
    const indexInput = document.getElementById('indexInput');
    const statusMessage = document.getElementById('statusMessage');
    const resultContainer = document.getElementById('resultContainer');

    // Standard columns that are NOT subjects
    // Convert to lowercase for easier matching
    const standardColumns = ['indexnumber', 'index number', 'index_number', 'id', 'name', 'student name', 'total', 'average', 'grade', 'status', 'rank'];

    searchBtn.addEventListener('click', performSearch);
    indexInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            performSearch();
        }
    });

    async function performSearch() {
        const indexNumber = indexInput.value.trim();
        
        if (!indexNumber) {
            showStatus('Please enter an Index Number.', 'error');
            return;
        }

        showStatus('Searching...', 'loading');
        resultContainer.classList.add('hidden');

        try {
            // Fetch the Excel file from the repository
            // Added cache-busting timestamp to ensure the browser always gets the latest file!
            const cacheBuster = new Date().getTime();
            const response = await fetch(`results.xlsx?t=${cacheBuster}`);
            
            if (!response.ok) {
                throw new Error('Could not find the results data file. Make sure results.xlsx is uploaded.');
            }

            const arrayBuffer = await response.arrayBuffer();
            
            // Parse Excel using SheetJS
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            
            // Assuming data is in the first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert sheet to JSON (array of objects)
            const data = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

            // Find the student
            // We'll search across all keys to find the one that looks like an index number
            let student = null;
            let indexKey = null;

            if (data.length > 0) {
                // Find the key that represents the index number
                const keys = Object.keys(data[0]);
                indexKey = keys.find(k => ['indexnumber', 'index number', 'id'].includes(k.toLowerCase().replace(/[^a-z]/g, '')));
                
                if (indexKey) {
                    // Find student where the index matches (case insensitive string comparison)
                    student = data.find(row => String(row[indexKey]).toLowerCase() === String(indexNumber).toLowerCase());
                }
            }

            if (student) {
                displayResult(student);
                showStatus('');
            } else {
                showStatus(`No results found for Index Number: ${indexNumber}`, 'error');
            }

        } catch (error) {
            console.error('Error fetching or parsing results:', error);
            showStatus(error.message || 'An error occurred while fetching results.', 'error');
        }
    }

    function showStatus(message, type = '') {
        statusMessage.textContent = message;
        statusMessage.className = 'status-message';
        if (type) {
            statusMessage.classList.add(type);
        }
    }

    function displayResult(student) {
        // Extract student info
        const keys = Object.keys(student);
        
        // Find specific keys dynamically
        const nameKey = keys.find(k => k.toLowerCase().includes('name'));
        const totalKey = keys.find(k => k.toLowerCase() === 'total' || k.toLowerCase().includes('total marks'));
        const averageKey = keys.find(k => k.toLowerCase().includes('average'));
        const gradeKey = keys.find(k => k.toLowerCase() === 'grade' || k.toLowerCase() === 'result');

        const name = nameKey ? student[nameKey] : 'Unknown Student';
        const id = indexInput.value.trim();
        const total = totalKey ? student[totalKey] : '-';
        const average = averageKey ? student[averageKey] : '-';
        const grade = gradeKey ? student[gradeKey] : '-';

        // Filter out subjects
        const subjects = [];
        keys.forEach(key => {
            const normalizedKey = key.toLowerCase().trim().replace(/[^a-z0-9]/g, '');
            // If the column name is not in our standard columns list, treat it as a subject
            const isStandard = standardColumns.some(std => normalizedKey.includes(std.replace(/[^a-z0-9]/g, '')));
            
            if (!isStandard && key.trim() !== '') {
                subjects.push({
                    name: key,
                    mark: student[key]
                });
            }
        });

        // Generate Subject HTML
        let subjectsHTML = '';
        if (subjects.length > 0) {
            subjects.forEach(sub => {
                subjectsHTML += `
                    <div class="subject-card">
                        <div class="subject-name">${sub.name}</div>
                        <div class="subject-mark">${sub.mark}</div>
                    </div>
                `;
            });
        } else {
            subjectsHTML = `<p style="grid-column: 1/-1; color: var(--text-muted);">No subject marks found.</p>`;
        }

        // Generate Full HTML
        const html = `
            <div class="result-card">
                <div class="student-info">
                    <div class="student-avatar">
                        <i class="fa-solid fa-user-graduate"></i>
                    </div>
                    <h2 class="student-name">${name}</h2>
                    <div class="student-id">
                        <i class="fa-solid fa-id-card"></i> ${id}
                    </div>
                </div>

                <div class="marks-container">
                    <h3>Subject Marks</h3>
                    <div class="subjects-grid">
                        ${subjectsHTML}
                    </div>
                </div>

                ${(total !== '-' || average !== '-' || grade !== '-') ? `
                <div class="summary-grid">
                    ${total !== '-' ? `
                    <div class="summary-item">
                        <div class="summary-label">Total</div>
                        <div class="summary-value value-total">${total}</div>
                    </div>` : ''}
                    
                    ${average !== '-' ? `
                    <div class="summary-item">
                        <div class="summary-label">Average</div>
                        <div class="summary-value value-average">${average}</div>
                    </div>` : ''}
                    
                    ${grade !== '-' ? `
                    <div class="summary-item">
                        <div class="summary-label">Grade</div>
                        <div class="summary-value value-grade">${grade}</div>
                    </div>` : ''}
                </div>
                ` : ''}
            </div>
            
            <div class="action-buttons" style="text-align: center; margin-top: 2rem;">
                <button id="downloadPdfBtn" class="download-btn">
                    <i class="fa-solid fa-file-pdf"></i> Download Result as PDF
                </button>
            </div>
        `;

        resultContainer.innerHTML = html;
        resultContainer.classList.remove('hidden');

        // Add event listener for PDF download
        document.getElementById('downloadPdfBtn').addEventListener('click', () => {
            // Populate the hidden PDF template
            document.getElementById('pdfName').textContent = name;
            document.getElementById('pdfIndex').textContent = id;
            
            const pdfTableBody = document.getElementById('pdfTableBody');
            pdfTableBody.innerHTML = '';
            
            if (subjects.length > 0) {
                subjects.forEach(sub => {
                    pdfTableBody.innerHTML += `
                        <tr>
                            <td style="border: 1px solid #000; padding: 12px; text-align: left;">${sub.name}</td>
                            <td style="border: 1px solid #000; padding: 12px; text-align: center; font-weight: bold;">${sub.mark}</td>
                        </tr>
                    `;
                });
            } else {
                pdfTableBody.innerHTML = `<tr><td colspan="2" style="border: 1px solid #000; padding: 12px; text-align: center;">No subject marks found</td></tr>`;
            }

            const pdfSummary = document.getElementById('pdfSummary');
            pdfSummary.innerHTML = '';
            if (total !== '-') pdfSummary.innerHTML += `<div>Total: ${total}</div>`;
            if (average !== '-') pdfSummary.innerHTML += `<div>Average: ${average}</div>`;
            if (grade !== '-') pdfSummary.innerHTML += `<div>Grade: ${grade}</div>`;

            const element = document.getElementById('pdfTemplate');
            
            // Temporarily show it for html2pdf to capture it
            element.style.display = 'block';

            const opt = {
                margin:       [20, 20, 20, 20],
                filename:     `${id}_Full_Result_Sheet.pdf`,
                image:        { type: 'jpeg', quality: 1.0 },
                html2canvas:  { 
                    scale: 2,
                    useCORS: true,
                    logging: false
                },
                jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
            };

            // Change button text to show loading
            const btn = document.getElementById('downloadPdfBtn');
            const originalText = btn.innerHTML;
            btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Generating PDF...';
            btn.disabled = true;

            html2pdf().set(opt).from(element).save().then(() => {
                element.style.display = 'none'; // Hide it again
                btn.innerHTML = originalText;
                btn.disabled = false;
            }).catch(err => {
                console.error("PDF Generation Error", err);
                element.style.display = 'none'; // Hide it again
                btn.innerHTML = originalText;
                btn.disabled = false;
            });
        });
    }
});
