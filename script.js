document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excel-file');
    const fileNameDisplay = document.getElementById('file-name');
    const filterSection = document.getElementById('filter-section');
    const classFiltersContainer = document.getElementById('class-filters');
    const teacherFiltersContainer = document.getElementById('teacher-filters');
    const locationFiltersContainer = document.getElementById('location-filters');
    const helpBtn = document.getElementById('help-btn');
    const helpModal = document.getElementById('help-modal');
    const closeHelpBtn = document.getElementById('close-help');
    const calendarContainer = document.getElementById('calendar-container');
    const clearFiltersBtn = document.getElementById('clear-filters-btn');
    const loadingOverlay = document.getElementById('loading');
    const filterSearchInput = document.getElementById('filter-search');
    const vakSearchInput = document.getElementById('vak-search');
    const vakSuggestionsContainer = document.getElementById('vak-suggestions');
    const vakFiltersContainer = document.getElementById('vak-filters');

    // Interne lijst van docenten (afkorting → volledige naam).
    // Gebruikt als fallback wanneer het Excel-bestand geen tweede blad bevat;
    // gegevens uit een meegeleverd "Docenten"-blad krijgen voorrang.
    const BUILT_IN_TEACHERS_MAP = {
        'ALL': 'Allard Nathalie',
        'BAEJ': 'Baert Jonas',
        'BEYN': 'Beyers Aäron',
        'BINNC': 'Binnemans Charlotte',
        'BOMI': 'Boeynaems Michael',
        'BOH': 'Bollaert Hiram',
        'BUN': 'Bungeneers Tom',
        'CASK': 'Casal Mosteiro Kelly',
        'CASW': 'Casteels Wim',
        'CHSV': 'Charleer Sven',
        'CHST': 'Charpentier Steven',
        'CLSA': 'Claes Sara',
        'COMRO': 'Cominotto Robin',
        'COPCH': 'Cop Christophe',
        'CROBR': 'Crols Bram',
        'DAT': 'Dams Tim',
        'DBM': 'De Bie Marijke',
        'DELH': 'De Clerck Liesbeth',
        'DDB': 'De Decker Benny',
        'DEDA': 'De Deken Dany',
        'DMAGE': 'de Magtige Ernie',
        'DEMP': 'De Meersman Patrick',
        'DEPM': 'De Pooter Marijn',
        'DEWA': 'De Ridder Ward',
        'DESNI': 'De Schutter Nick',
        'DEKO': 'De Smet Koen',
        'VOJE': 'De Vos Jeroen',
        'DEVT': 'De Vries Thomas',
        'DEVU': 'De Vuyst Marijke',
        'WITE': 'De Wit Eveline',
        'DEBE': 'Debeuf Ann',
        'DELA': 'Delafortry Lina',
        'DELIK': 'Delien Kim',
        'DIEP': 'Diependaele Kevin',
        'DOST': 'Doggen Stijn',
        'DRFR': 'Draulans Frederik',
        'DRE': 'Dreessen Kristel',
        'DRUS': 'Druyts Stijn',
        'EHNA': 'El Hajjouti Nabil',
        'FRA': 'Frateur Wouter',
        'GEJI': 'Geens Jill',
        'GEM': 'Gemoets Ann',
        'GHAN': 'Ghanoudi Hakim',
        'GOEL': 'Goedemé Lien',
        'GOC': 'Goidts Christophe',
        'GOM': 'Gombeer Wim',
        'GOJO': 'Gossé Joke',
        'GUJ': 'Guldentops Jan',
        'HADD': 'Haddouchi Hassan',
        'HAZ': 'Hazebroek Maarten',
        'HEIR': 'Heirbaut Steven',
        'HESA': 'Hendrickx Sanne',
        'HERB': 'Herman Bruno',
        'HOSE': 'Horsmans Serge',
        'HYE': 'Hye Dirk',
        'JAMY': 'Janse Myrthe',
        'JADE': 'Janssens Dennis',
        'JOOI': 'Joosen Isabelle',
        'KOBR': 'Kortleven Bram',
        'LEFL': 'Lefebure Laurens',
        'LEEL': 'Leonard Elias',
        'LET': 'Lettany Barbara',
        'LIE': 'Listhaeghe Erwin',
        'LIV': 'Livens Wim',
        'LUP': 'Luyckx Philip',
        'LUMA': 'Luyts Maarten',
        'MAND': 'Manderyck Isaac',
        'MAF': 'Marain Frederik',
        'MASV': 'Mariën Sven',
        'MAKR': 'Martens Kristoff',
        'MASS': 'Masset Yves',
        'MES': 'Mermans Sandra',
        'MICK': 'Michiels Kristof',
        'MOW': 'Moons Wouter',
        'NYV': 'Nys Vincent',
        'ORL': 'Orlent Indra',
        'OVBR': 'Overstijns Bram',
        'PAE': 'Paesmans Mietje',
        'PNI': 'Pauwels Nick',
        'PEWO': 'Peetermans Wouter',
        'PEA': 'Peeters Andreas',
        'PEEJA': 'Peeters Jannes',
        'PELAN': 'Peeters Lander',
        'PET': 'Peeters Tom',
        'POS': 'Possemiers Philippe',
        'PYP': 'Pype Lynn',
        'ROI': 'Robijns Iris',
        'RODI': 'Roelants Dieter',
        'ROMM': 'Rolus Magalie',
        'ROSM': 'Rosseau Marc',
        'RUB': 'Rubberecht Pieter',
        'SCMAR': 'Schraepen Mario',
        'SCRPH': 'Schraepen Philippe',
        'SERA': 'Serneels Alexander',
        'SIAN': 'Similon Andie',
        'SPPE': 'Spaas Peter',
        'SPO': 'Spittaels Olaf',
        'STEW': 'Stevens Wouter',
        'STT': 'Stoops Tim',
        'STD': 'Sturm Dimitri',
        'TSA': 'Thijs Alain',
        'TIRU': 'Tiebos Ruud',
        'TIEPT': 'Tiepermann Tom',
        'TILD': 'Tillemans David',
        'VAN': 'Van Acker Nick',
        'VAWI': 'Van Aerschot Wim',
        'VAV': 'Van Assche Veerle',
        'VBSA': 'Van Battel Sam',
        'VASE': 'Van Camp Serge',
        'CAMV': 'Van Camp Vincent',
        'VVI': 'Van de Venne Ingeborg',
        'VBH': 'Van den Bulck Helga',
        'VDHS': 'Van den Heuvel Sylvia',
        'VDL': 'Van Den Langenbergh Maïté',
        'POEJ': 'Van den Poel Jan',
        'VAHA': 'Van der Kraan Harold',
        'VDES': 'Van der Meulen Esthel',
        'DIJM': 'van Dijk Maria',
        'VAEY': 'Van Eyken Koen',
        'VAGI': 'Van Gils Christel',
        'VGG': 'Van Grieken Geert',
        'VABA': 'Van Hecke Bavo',
        'LOOE': 'Van Loo Erwin',
        'VLO': 'Van Looveren Ilse',
        'VADK': 'Van Merode Dirk',
        'OVMA': 'Van Overveldt Maarten',
        'VRJO': 'Van Riel Jonas',
        'VRSTE': 'Van Rossem Stephane',
        'VANTB': 'Van Thielen Bart',
        'VAM': 'van Varik Meeuwis',
        'VDVP': 'Vander Vennet Pieter',
        'HULG': 'Vanhulle Geert',
        'VERHL': 'Verhaert Loes',
        'VEDD': 'Verhulst David',
        'VETH': 'Verhuyck Thomas',
        'VEK': 'Verlinden Karen',
        'VEDAD': 'Vermonden David',
        'VRB': 'Verstraeten Bartholomeus',
        'VIS': 'Vissers Nadia',
        'WAEDI': 'Waeyaert Dimitri',
        'WEU': 'Weuts Anne-Elisabeth',
        'WIOL': 'Wille Olivier',
        'WIV': 'Willemen Vanessa/Vaya',
        'YPS': 'Yperzeele Saskia'
    };

    let rawSchedule = [];
    let teachersMap = { ...BUILT_IN_TEACHERS_MAP };
    let allClasses = new Set();
    let selectedClasses = new Set();
    let allTeachers = new Set();
    let selectedTeachers = new Set();
    let allLocations = new Set();
    let selectedLocations = new Set();
    let allVaks = new Set();
    let selectedVaks = new Set();

    fileInput.addEventListener('change', handleFileUpload);

    if (helpBtn && helpModal && closeHelpBtn) {
        helpBtn.addEventListener('click', () => {
            helpModal.classList.remove('hidden');
        });

        closeHelpBtn.addEventListener('click', () => {
            helpModal.classList.add('hidden');
        });

        helpModal.addEventListener('click', (e) => {
            if (e.target === helpModal) {
                helpModal.classList.add('hidden');
            }
        });
    }

    clearFiltersBtn.addEventListener('click', () => {
        selectedClasses.clear();
        selectedTeachers.clear();
        selectedLocations.clear();
        selectedVaks.clear();
        document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
        renderVakChips();
        if (vakSearchInput) vakSearchInput.value = '';
        hideVakSuggestions();
        renderCalendar();
    });

    if (filterSearchInput) {
        filterSearchInput.addEventListener('input', applySearchFilter);
    }

    if (vakSearchInput) {
        vakSearchInput.addEventListener('input', showVakSuggestions);
        vakSearchInput.addEventListener('focus', () => {
            if (vakSearchInput.value.trim()) showVakSuggestions();
        });
        vakSearchInput.addEventListener('blur', () => {
            // Delay so mousedown on suggestion can fire first
            setTimeout(hideVakSuggestions, 150);
        });
        vakSearchInput.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                vakSearchInput.value = '';
                hideVakSuggestions();
            }
        });
    }

    function applySearchFilter() {
        const query = (filterSearchInput.value || '').trim().toLowerCase();
        const groups = document.querySelectorAll('#filter-section .filter-group');
        groups.forEach(group => {
            // Vak-groep heeft eigen zoekveld; globale search niet toepassen
            if (group.querySelector('.vak-filter-body')) return;
            const chips = group.querySelectorAll('.chip');
            let visibleCount = 0;
            chips.forEach(chip => {
                const haystack = (chip.textContent + ' ' + (chip.title || '')).toLowerCase();
                const match = !query || haystack.includes(query);
                chip.classList.toggle('hidden-by-search', !match);
                if (match) visibleCount++;
            });
            group.classList.toggle('no-search-matches', query && visibleCount === 0);
            // Auto-expand groups when searching so matches are visible
            if (query && visibleCount > 0) {
                group.classList.remove('collapsed');
            }
        });
    }

    function renderVakChips() {
        if (!vakFiltersContainer) return;
        vakFiltersContainer.innerHTML = '';
        Array.from(selectedVaks).sort().forEach(vak => {
            const chip = document.createElement('div');
            chip.className = 'chip active';
            chip.textContent = vak;
            chip.title = 'Klik om te verwijderen';
            chip.addEventListener('click', () => {
                selectedVaks.delete(vak);
                renderVakChips();
                renderCalendar();
                if (vakSearchInput && vakSearchInput.value.trim()) showVakSuggestions();
            });
            vakFiltersContainer.appendChild(chip);
        });
    }

    function showVakSuggestions() {
        if (!vakSuggestionsContainer || !vakSearchInput) return;
        const query = vakSearchInput.value.trim().toLowerCase();
        if (!query) { hideVakSuggestions(); return; }

        const matches = Array.from(allVaks)
            .filter(v => !selectedVaks.has(v))
            .filter(v => v.toLowerCase().includes(query))
            .sort();

        vakSuggestionsContainer.innerHTML = '';
        if (matches.length === 0) {
            const empty = document.createElement('div');
            empty.className = 'vak-suggestion disabled';
            empty.textContent = 'Geen vakken gevonden';
            vakSuggestionsContainer.appendChild(empty);
        } else {
            const MAX = 15;
            matches.slice(0, MAX).forEach(vak => {
                const item = document.createElement('div');
                item.className = 'vak-suggestion';
                item.textContent = vak;
                // mousedown vuurt vóór blur, zodat de suggestie niet verdwijnt voor de klik
                item.addEventListener('mousedown', (e) => {
                    e.preventDefault();
                    selectedVaks.add(vak);
                    vakSearchInput.value = '';
                    renderVakChips();
                    hideVakSuggestions();
                    renderCalendar();
                    vakSearchInput.focus();
                });
                vakSuggestionsContainer.appendChild(item);
            });
            if (matches.length > MAX) {
                const more = document.createElement('div');
                more.className = 'vak-suggestions-more';
                more.textContent = `...en nog ${matches.length - MAX} meer (verfijn zoekterm)`;
                vakSuggestionsContainer.appendChild(more);
            }
        }
        vakSuggestionsContainer.classList.remove('hidden');
    }

    function hideVakSuggestions() {
        if (vakSuggestionsContainer) vakSuggestionsContainer.classList.add('hidden');
    }

    function handleFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        fileNameDisplay.textContent = file.name;
        loadingOverlay.classList.remove('hidden');

        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            try {
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });

                // Process Sheet 1 (Schedule)
                const firstSheetName = workbook.SheetNames[0];
                const scheduleSheet = workbook.Sheets[firstSheetName];
                // Using raw: true gets us Date objects instead of randomly formatted strings
                rawSchedule = XLSX.utils.sheet_to_json(scheduleSheet, { defval: "", raw: true });

                // Start van de interne lijst; Sheet 2 (indien aanwezig) overschrijft/aanvult
                teachersMap = { ...BUILT_IN_TEACHERS_MAP };

                if (workbook.SheetNames.length > 1) {
                    const secondSheetName = workbook.SheetNames[1];
                    const teacherSheet = workbook.Sheets[secondSheetName];
                    const teacherData = XLSX.utils.sheet_to_json(teacherSheet, { header: 1, defval: "" });
                    // Assuming row 0 is header. Kolom 1 = Naam, Kolom 2 = Afkorting
                    for (let i = 1; i < teacherData.length; i++) {
                        const row = teacherData[i];
                        if (row && row.length >= 2) {
                            const naam = row[0].toString().trim();
                            const afkort = row[1].toString().trim();
                            if (afkort && naam) {
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
        allTeachers.clear();
        selectedTeachers.clear();
        allLocations.clear();
        selectedLocations.clear();
        allVaks.clear();
        selectedVaks.clear();

        rawSchedule.forEach(row => {
            const klassenStr = row['Klas(sen)'];
            if (klassenStr) {
                const klassenArr = klassenStr.split(',').map(k => k.trim()).filter(k => k);
                klassenArr.forEach(k => allClasses.add(k));
            }

            const docentAfkort = row['Docent'];
            if (docentAfkort) {
                const docentArr = docentAfkort.toString().split(',').map(d => d.trim()).filter(d => d);
                docentArr.forEach(d => allTeachers.add(d));
            }

            const lokalenStr = row['Lokalen'];
            if (lokalenStr) {
                const lokalenArr = lokalenStr.toString().split(',').map(l => l.trim()).filter(l => l);
                lokalenArr.forEach(l => allLocations.add(l));
            }

            const vakStr = row['Vak'];
            if (vakStr) {
                const v = vakStr.toString().trim();
                if (v) allVaks.add(v);
            }
        });

        // Sort classes alphabetically
        const sortedClasses = Array.from(allClasses).sort();
        // Sort teachers alphabetically
        const sortedTeachers = Array.from(allTeachers).sort();
        // Sort locations alphabetically
        const sortedLocations = Array.from(allLocations).sort();

        // Build Class Filter UI
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

        // Build Teacher Filter UI
        teacherFiltersContainer.innerHTML = '';
        sortedTeachers.forEach(docentAfkort => {
            const chip = document.createElement('div');
            chip.className = 'chip';
            const fullNaam = teachersMap[docentAfkort];
            // Toon de afkorting; volledige naam (indien gekend) verschijnt als tooltip op hover
            chip.textContent = docentAfkort;
            if (fullNaam && fullNaam !== docentAfkort) {
                chip.title = fullNaam;
            }
            chip.addEventListener('click', () => {
                chip.classList.toggle('active');
                if (selectedTeachers.has(docentAfkort)) {
                    selectedTeachers.delete(docentAfkort);
                } else {
                    selectedTeachers.add(docentAfkort);
                }
                renderCalendar();
            });
            teacherFiltersContainer.appendChild(chip);
        });

        // Build Location Filter UI
        locationFiltersContainer.innerHTML = '';
        sortedLocations.forEach(loc => {
            const chip = document.createElement('div');
            chip.className = 'chip';
            chip.textContent = loc;
            chip.addEventListener('click', () => {
                chip.classList.toggle('active');
                if (selectedLocations.has(loc)) {
                    selectedLocations.delete(loc);
                } else {
                    selectedLocations.add(loc);
                }
                renderCalendar();
            });
            locationFiltersContainer.appendChild(chip);
        });

        // Reset vak-UI (typeahead + chips)
        if (vakSearchInput) vakSearchInput.value = '';
        hideVakSuggestions();
        renderVakChips();

        // Re-apply any active search query to the freshly built chips
        applySearchFilter();

        // Render everything initially
        renderCalendar();
    }

    // Helper to parse dates safely (handles both SheetJS Date objects and DD/MM/YYYY strings)
    function parseDateObj(dateVal) {
        if (!dateVal) return null;

        // If SheetJS already parsed it as a Date object (cellDates: true)
        if (dateVal instanceof Date) {
            // SheetJS dates from Excel represent the local date accurately 
            // when using local time methods (.getFullYear(), .getMonth(), .getDate())
            return new Date(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate());
        }

        // Check for common DD/MM/YYYY string pattern from CSV
        if (typeof dateVal === 'string') {
            const parts = dateVal.match(/(\d+)\/(\d+)\/(\d+)/);
            if (parts) {
                let day = parseInt(parts[1], 10);
                let month = parseInt(parts[2], 10) - 1; // 0-indexed
                let year = parseInt(parts[3], 10);
                if (year < 100) year += 2000;
                return new Date(year, month, day);
            }

            // Fallback for native formats
            const d = new Date(dateVal);
            return isNaN(d.getTime()) ? null : d;
        }

        return null;
    }

    function renderCalendar() {
        let filtered = rawSchedule.filter(row => {
            let classMatch = true;
            if (selectedClasses.size > 0) {
                const klassenStr = row['Klas(sen)'] || '';
                const klassenArr = klassenStr.split(',').map(k => k.trim());
                classMatch = klassenArr.some(k => selectedClasses.has(k));
            }

            let teacherMatch = true;
            if (selectedTeachers.size > 0) {
                const docentAfkort = row['Docent'] || '';
                const docentArr = docentAfkort.toString().split(',').map(d => d.trim());
                teacherMatch = docentArr.some(d => selectedTeachers.has(d));
            }

            let locationMatch = true;
            if (selectedLocations.size > 0) {
                const lokalenStr = row['Lokalen'] || '';
                const lokalenArr = lokalenStr.toString().split(',').map(l => l.trim());
                locationMatch = lokalenArr.some(l => selectedLocations.has(l));
            }

            let vakMatch = true;
            if (selectedVaks.size > 0) {
                const vak = (row['Vak'] || '').toString().trim();
                vakMatch = selectedVaks.has(vak);
            }

            return classMatch && teacherMatch && locationMatch && vakMatch;
        });

        const dayNames = ['Zondag', 'Maandag', 'Dinsdag', 'Woensdag', 'Donderdag', 'Vrijdag', 'Zaterdag'];

        // Enrich objects with parsed Date and Time
        const enriched = filtered.map(row => {
            const dateObj = parseDateObj(row['Datum']);
            return {
                ...row,
                _dateObj: dateObj,
                _timestamp: dateObj ? dateObj.getTime() : 0,
                _beginMinutes: parseTime(row['Begin']),
                _formattedBegin: formatTimeField(row['Begin']),
                _formattedEindtijd: formatTimeField(row['Eindtijd'])
            };
        }).filter(r => r._dateObj !== null);

        // Sort by Date, then Begin time
        enriched.sort((a, b) => {
            if (a._timestamp !== b._timestamp) return a._timestamp - b._timestamp;
            return a._beginMinutes - b._beginMinutes;
        });

        // Merge duplicate moments (same date/time/subject/classes/location) — combine docenten
        const merged = mergeDuplicateExams(enriched);

        if (merged.length === 0) {
            calendarContainer.innerHTML = '<p style="text-align: center; color: var(--text-secondary); padding: 2rem;">Geen examens gevonden voor deze selectie.</p>';
            return;
        }

        // Bepaal de vroegste datum over alle examens heen voor consistente weektelling
        let globalEarliestDate = null;
        rawSchedule.forEach(row => {
            const d = parseDateObj(row['Datum']);
            if (d && (!globalEarliestDate || d.getTime() < globalEarliestDate.getTime())) {
                globalEarliestDate = d;
            }
        });

        const earliestDate = globalEarliestDate || merged[0]._dateObj;
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
        merged.forEach(exam => {
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
                    const docentList = docentAfkort.toString().split(',').map(d => d.trim()).filter(d => d);
                    const fullDocent = docentList
                        .map(d => teachersMap[d] && teachersMap[d] !== d ? `${teachersMap[d]} (${d})` : d)
                        .join(', ');

                    html += `
                            <div class="exam-card" onclick="this.classList.toggle('expanded')">
                                <div class="card-header">
                                    <div class="exam-time">
                                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>
                                        ${exam._formattedBegin || '-'} - ${exam._formattedEindtijd || '-'} 
                                        ${exam['Uren'] ? `<span style="color:var(--text-secondary); font-size:0.75rem; margin-left:auto;">(${exam['Uren']})</span>` : ''}
                                    </div>
                                    <div class="exam-subject">${exam['Vak'] || 'Onbekend Vak'}</div>
                                    
                                    <div class="exam-detail" title="Docent">
                                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg>
                                        <span>${fullDocent || '-'}</span>
                                    </div>
                                    <div class="expand-icon">
                                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>
                                    </div>
                                </div>
                                <div class="card-body">
                                    <div class="exam-detail" title="Locatie">
                                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path><circle cx="12" cy="10" r="3"></circle></svg>
                                        <span class="word-wrap">${(exam['Lokalen'] || '-').split(',').join(', ')}</span>
                                    </div>
                                    <div class="exam-detail" title="Klassen">
                                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
                                        <span class="word-wrap">${(exam['Klas(sen)'] || '-').split(',').join(', ')}</span>
                                    </div>
                                    ${exam['Lestekst'] ? `
                                    <div class="exam-detail" style="color: #cbd5e1; font-style: italic; margin-top: 4px;">
                                        <span class="word-wrap">${exam['Lestekst']}</span>
                                    </div>` : ''}
                                </div>
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

    function mergeDuplicateExams(exams) {
        const groups = new Map();
        exams.forEach(exam => {
            const key = [
                exam._timestamp,
                exam._formattedBegin,
                exam._formattedEindtijd,
                (exam['Vak'] || '').toString().trim(),
                (exam['Klas(sen)'] || '').toString().trim(),
                (exam['Lokalen'] || '').toString().trim(),
                (exam['Uren'] || '').toString().trim(),
                (exam['Lestekst'] || '').toString().trim()
            ].join('|');

            if (!groups.has(key)) {
                groups.set(key, { exam: { ...exam }, docenten: [] });
            }
            const group = groups.get(key);
            (exam['Docent'] || '').toString().split(',').map(d => d.trim()).filter(d => d).forEach(d => {
                if (!group.docenten.includes(d)) group.docenten.push(d);
            });
        });

        return Array.from(groups.values()).map(g => {
            g.exam['Docent'] = g.docenten.join(', ');
            return g.exam;
        });
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

    function formatTimeField(val) {
        if (!val) return '';
        if (val instanceof Date) {
            // SheetJS times-only (without date) are mapped such that UTC hours match the text
            return `${val.getUTCHours()}:${val.getUTCMinutes().toString().padStart(2, '0')}`;
        }
        return val.toString();
    }

    function parseTime(timeVal) {
        if (!timeVal) return 0;
        if (timeVal instanceof Date) {
            return timeVal.getUTCHours() * 60 + timeVal.getUTCMinutes();
        }
        const parts = timeVal.toString().split(':');
        if (parts.length >= 2) {
            return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
        }
        return 0;
    }
});
