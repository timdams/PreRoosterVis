document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excel-file');
    const fileNameDisplay = document.getElementById('file-name');
    const filterSection = document.getElementById('filter-section');
    const classFiltersContainer = document.getElementById('class-filters');
    const calendarContainer = document.getElementById('calendar-container');
    const clearFiltersBtn = document.getElementById('clear-filters-btn');
    const loadingOverlay = document.getElementById('loading');

    let rawSchedule = [];
    let teachersMap = {};
    let allClasses = new Set();
    let selectedClasses = new Set();

    fileInput.addEventListener('change', handleFileUpload);
    clearFiltersBtn.addEventListener('click', () => {
        selectedClasses.clear();
        document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
        renderCalendar();
    });

    function handleFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        fileNameDisplay.textContent = file.name;
        loadingOverlay.classList.remove('hidden');

        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            try {
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Process Sheet 1 (Schedule)
                const firstSheetName = workbook.SheetNames[0];
                const scheduleSheet = workbook.Sheets[firstSheetName];
                // Using raw: false ensures dates and times are formatted as strings based on Excel format.
                rawSchedule = XLSX.utils.sheet_to_json(scheduleSheet, { defval: "", raw: false });

                // Process Sheet 2 (Teachers)
                if (workbook.SheetNames.length > 1) {
                    const secondSheetName = workbook.SheetNames[1];
                    const teacherSheet = workbook.Sheets[secondSheetName];
                    const teacherData = XLSX.utils.sheet_to_json(teacherSheet, { header: 1, defval: "" });
                    // Assuming row 0 is header. Kolom 1 = Naam, Kolom 2 = Afkorting
                    teachersMap = {};
                    for (let i = 1; i < teacherData.length; i++) {
                        const row = teacherData[i];
                        if (row && row.length >= 2) {
                            const naam = row[0].toString().trim();
                            const afkort = row[1].toString().trim();
                            if (afkort) {
                                teachersMap[afkort] = naam;
                            }
                        }
                    }
                }

                extractClassesAndInitialize();
                
                filterSection.classList.remove('hidden');
                calendarContainer.classList.remove('hidden');
            } catch (error) {
                console.error("Error reading file:", error);
                alert("Fout bij inladen van het bestand. Zorg dat het een geldig Excel of CSV bestand is.");
            } finally {
                loadingOverlay.classList.add('hidden');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function extractClassesAndInitialize() {
        allClasses.clear();
        selectedClasses.clear();
        
        rawSchedule.forEach(row => {
            const klassenStr = row['Klas(sen)'];
            if (klassenStr) {
                const klassenArr = klassenStr.split(',').map(k => k.trim()).filter(k => k);
                klassenArr.forEach(k => allClasses.add(k));
            }
        });

        // Sort classes alphabetically
        const sortedClasses = Array.from(allClasses).sort();
        
        // Build Filter UI
        classFiltersContainer.innerHTML = '';
        sortedClasses.forEach(cls => {
            const chip = document.createElement('div');
            chip.className = 'chip';
            chip.textContent = cls;
            chip.addEventListener('click', () => {
                chip.classList.toggle('active');
                if (selectedClasses.has(cls)) {
                    selectedClasses.delete(cls);
                } else {
                    selectedClasses.add(cls);
                }
                renderCalendar();
            });
            classFiltersContainer.appendChild(chip);
        });

        // Render everything initially
        renderCalendar();
    }

    // Helper to parse "DD/MM/YYYY" to Date object safely
    function parseDateStr(dateStr) {
        if (!dateStr) return null;
        
        // Check for common DD/MM/YYYY pattern
        const parts = dateStr.match(/(\d+)\/(\d+)\/(\d+)/);
        if (parts) {
            let day = parseInt(parts[1], 10);
            let month = parseInt(parts[2], 10) - 1; // 0-indexed
            let year = parseInt(parts[3], 10);
            if (year < 100) year += 2000;
            return new Date(year, month, day);
        }
        
        // Fallback for native formats
        const d = new Date(dateStr);
        return isNaN(d.getTime()) ? null : d;
    }

    function renderCalendar() {
        let filtered = rawSchedule.filter(row => {
            if (selectedClasses.size === 0) return true; // Show all if nothing selected
            
            const klassenStr = row['Klas(sen)'] || '';
            const klassenArr = klassenStr.split(',').map(k => k.trim());
            return klassenArr.some(k => selectedClasses.has(k));
        });

        const dayNames = ['Zondag', 'Maandag', 'Dinsdag', 'Woensdag', 'Donderdag', 'Vrijdag', 'Zaterdag'];
        
        // Enrich objects with parsed Date and Time
        const enriched = filtered.map(row => {
            const dateObj = parseDateStr(row['Datum']);
            return {
                ...row,
                _dateObj: dateObj,
                _timestamp: dateObj ? dateObj.getTime() : 0,
                _beginMinutes: parseTime(row['Begin'])
            };
        }).filter(r => r._dateObj !== null);

        // Sort by Date, then Begin time
        enriched.sort((a, b) => {
            if (a._timestamp !== b._timestamp) return a._timestamp - b._timestamp;
            return a._beginMinutes - b._beginMinutes;
        });

        if (enriched.length === 0) {
            calendarContainer.innerHTML = '<p style="text-align: center; color: var(--text-secondary); padding: 2rem;">Geen examens gevonden voor deze selectie.</p>';
            return;
        }

        const earliestDate = enriched[0]._dateObj;
        // Fallback: earliest Monday
        const startMonday = new Date(earliestDate);
        const dayOfWeek = startMonday.getDay();
        const diff = startMonday.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
        startMonday.setDate(diff);

        // Group into Week 1, Week 2, Week 3
        const weeks = [
            generateWeekGrid(startMonday, 0),
            generateWeekGrid(startMonday, 1),
            generateWeekGrid(startMonday, 2),
        ];

        // Assign exams to correct day & week
        enriched.forEach(exam => {
            const d = exam._dateObj;
            // Calculate difference in days safely without daylight saving time issues
            // By resetting time to midnight UTC
            const dUTC = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
            const startUTC = Date.UTC(startMonday.getFullYear(), startMonday.getMonth(), startMonday.getDate());
            
            const diffDays = Math.floor((dUTC - startUTC) / (1000 * 60 * 60 * 24));
            
            const weekIdx = Math.floor(diffDays / 7);
            const dayIdx = diffDays % 7;
            
            // Limit to Mon-Fri of the 3 weeks
            if (weekIdx >= 0 && weekIdx <= 2 && dayIdx >= 0 && dayIdx <= 4) {
                weeks[weekIdx].days[dayIdx].exams.push(exam);
            }
        });

        // Render HTML
        let html = '';
        weeks.forEach((week, i) => {
            html += `
                <div class="week-view glass-panel">
                    <div class="week-title">Week ${i + 1} &nbsp;<span style="font-size: 1rem; font-weight: 400;">(${formatDateShort(week.startDate)} - ${formatDateShort(addDays(week.startDate, 4))})</span></div>
                    <div class="days-grid">
            `;
            
            week.days.forEach(day => {
                html += `
                        <div class="day-column">
                            <div class="day-header">
                                ${dayNames[day.date.getDay()]}
                                <span class="date">${formatDateShort(day.date)}</span>
                            </div>
                `;
                
                if (day.exams.length === 0) {
                    html += `<div style="text-align: center; color: var(--card-border); padding: 1rem; font-size: 0.85rem;">Geen examens</div>`;
                }

                day.exams.forEach(exam => {
                    const docentAfkort = exam['Docent'] || '';
                    const fullDocent = teachersMap[docentAfkort] ? `${teachersMap[docentAfkort]} (${docentAfkort})` : docentAfkort;
                    
                    html += `
                            <div class="exam-card">
                                <div class="exam-time">
                                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>
                                    ${exam['Begin'] || '-'} - ${exam['Eindtijd'] || '-'} 
                                    ${exam['Uren'] ? `<span style="color:var(--text-secondary); font-size:0.75rem; margin-left:auto;">(${exam['Uren']})</span>` : ''}
                                </div>
                                <div class="exam-subject">${exam['Vak'] || 'Onbekend Vak'}</div>
                                
                                <div class="exam-detail" title="Docent">
                                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg>
                                    <span>${fullDocent || '-'}</span>
                                </div>
                                <div class="exam-detail" title="Locatie">
                                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path><circle cx="12" cy="10" r="3"></circle></svg>
                                    <span>${(exam['Lokalen'] || '-').split(',').join(', ')}</span>
                                </div>
                                <div class="exam-detail" title="Klassen">
                                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
                                    <span>${exam['Klas(sen)'] || '-'}</span>
                                </div>
                                ${exam['Lestekst'] ? `
                                <div class="exam-detail" style="color: #cbd5e1; font-style: italic; margin-top: 4px;">
                                    <span>${exam['Lestekst']}</span>
                                </div>` : ''}
                            </div>
                    `;
                });
                
                html += `
                        </div>
                `;
            });
            
            html += `
                    </div>
                </div>
            `;
        });

        calendarContainer.innerHTML = html;
    }

    function addDays(date, days) {
        const d = new Date(date);
        d.setDate(d.getDate() + days);
        return d;
    }

    function generateWeekGrid(startMonday, weekOffset) {
        const result = {
            startDate: addDays(startMonday, weekOffset * 7),
            days: []
        };
        for (let i = 0; i < 5; i++) { // Mon-Fri
            result.days.push({
                date: addDays(result.startDate, i),
                exams: []
            });
        }
        return result;
    }

    function formatDateShort(d) {
        const dd = String(d.getDate()).padStart(2, '0');
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        return `${dd}/${mm}`;
    }

    function parseTime(timeStr) {
        if (!timeStr) return 0;
        const parts = timeStr.toString().split(':');
        if (parts.length >= 2) {
            return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
        }
        return 0;
    }
});
