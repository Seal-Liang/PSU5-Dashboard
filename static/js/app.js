document.addEventListener('DOMContentLoaded', () => {
    // Navigation
    const navLinks = {
        'nav-dashboard': 'view-dashboard',
        'nav-upload': 'view-upload',
        'nav-settings': 'view-settings'
    };

    Object.keys(navLinks).forEach(navId => {
        document.getElementById(navId).addEventListener('click', (e) => {
            e.preventDefault();
            Object.keys(navLinks).forEach(id => document.getElementById(id).classList.remove('active'));
            e.target.classList.add('active');

            Object.values(navLinks).forEach(viewId => document.getElementById(viewId).style.display = 'none');
            document.getElementById(navLinks[navId]).style.display = 'block';
            
            if (navId === 'nav-dashboard') {
                document.getElementById('page-title').textContent = 'Overview Dashboard';
                document.getElementById('dashboard-filters').style.display = 'flex';
                const actions = document.querySelector('.header-actions');
                if(actions) actions.style.display = 'flex';
                if(globalDataCache) updateTimeUI(); 
            } else if (navId === 'nav-upload') {
                document.getElementById('page-title').textContent = 'Upload Data';
                document.getElementById('dashboard-filters').style.display = 'none';
                const actions = document.querySelector('.header-actions');
                if(actions) actions.style.display = 'none';
            } else if (navId === 'nav-settings') {
                document.getElementById('page-title').textContent = 'Settings';
                document.getElementById('dashboard-filters').style.display = 'none';
                const actions = document.querySelector('.header-actions');
                if(actions) actions.style.display = 'none';
            }
        });
    });

    let globalDataCache = null;
    let buChartInstance = null;
    let utilChartInstance = null;
    let pieChartInstance = null;
    
    let availablePeriods = { Year: [], Quarter: [], Month: [] };
    let currentPeriodIdx = { Year: 0, Quarter: 0, Month: 0 };

    window.downloadWidget = function(widgetId, name) {
        if (typeof html2canvas === 'undefined') return alert('Library not loaded. Please restart application.');
        const widget = document.getElementById(widgetId);
        const btns = widget.querySelectorAll('.export-btn');
        btns.forEach(b => b.style.display = 'none');
        html2canvas(widget, {backgroundColor: '#ffffff'}).then(canvas => {
            btns.forEach(b => b.style.display = '');
            const url = canvas.toDataURL('image/png', 1.0);
            const a = document.createElement('a');
            a.href = url; a.download = `${name}.png`; a.click();
        }).catch(err => {
            btns.forEach(b => b.style.display = '');
            console.error(err);
        });
    };
    
    window.copyWidget = function(widgetId) {
        if (typeof html2canvas === 'undefined') return alert('Library not loaded. Please restart application.');
        const widget = document.getElementById(widgetId);
        const btns = widget.querySelectorAll('.export-btn');
        btns.forEach(b => b.style.display = 'none');
        html2canvas(widget, {backgroundColor: '#ffffff'}).then(canvas => {
            btns.forEach(b => b.style.display = '');
            canvas.toBlob(blob => {
                if(!blob) return;
                navigator.clipboard.write([new ClipboardItem({'image/png': blob})])
                .then(()=>alert('Widget copied to clipboard!'))
                .catch(err => {
                    console.error(err);
                    alert("Failed to copy image to clipboard.");
                });
            }, 'image/png', 1.0);
        }).catch(err => {
            btns.forEach(b => b.style.display = '');
            console.error(err);
        });
    };

    function getAssignedMonthLabel(dateStr) {
        let formattedStr = dateStr.replace(/\//g, '-');
        let d = new Date(formattedStr);
        if(isNaN(d)) return "";
        let d2 = new Date(d);
        d2.setDate(d.getDate() + 7);
        let year = d.getFullYear();
        let month = d.getMonth() + 1;
        if (d.getMonth() !== d2.getMonth()) {
            month = d2.getMonth() + 1;
            year = d2.getFullYear();
        }
        return `${year}-${String(month).padStart(2, '0')}`;
    }

    function getAssignedQuarterLabel(dateStr) {
        let mLabel = getAssignedMonthLabel(dateStr);
        if(!mLabel) return "";
        let [y, m] = mLabel.split('-');
        let q = Math.ceil(parseInt(m) / 3);
        return `${y} Q${q}`;
    }

    function loadDataAndRender() {
        fetch('/api/data').then(res => res.json()).then(data => {
            globalDataCache = data;
            populateFilters(data);
            
            const selectedTeam = document.getElementById('settingsTeamSelect').value;
            if (selectedTeam) renderTeamLog(selectedTeam);
            
            renderHolidays();
            updateTimeUI();
        }).catch(console.error);
    }

    function populateFilters(data) {
        const teams = new Set();
        availablePeriods = { Year: [], Quarter: [], Month: [] };

        (data.WeeklyData || []).forEach(wk => {
            (wk.records || []).forEach(r => teams.add(r.department));
            let mLabel = getAssignedMonthLabel(wk.date);
            let qLabel = getAssignedQuarterLabel(wk.date);
            let yLabel = mLabel.split('-')[0];
            if (yLabel) {
                availablePeriods.Year.push(yLabel);
                availablePeriods.Quarter.push(qLabel);
                availablePeriods.Month.push(mLabel);
            }
        });

        availablePeriods.Year = [...new Set(availablePeriods.Year)].sort();
        availablePeriods.Quarter = [...new Set(availablePeriods.Quarter)].sort();
        availablePeriods.Month = [...new Set(availablePeriods.Month)].sort();
        
        currentPeriodIdx.Year = Math.max(0, availablePeriods.Year.length - 1);
        currentPeriodIdx.Quarter = Math.max(0, availablePeriods.Quarter.length - 1);
        currentPeriodIdx.Month = Math.max(0, availablePeriods.Month.length - 1);

        const teamArr = Array.from(teams).sort();
        
        const dashFilter = document.getElementById('teamFilter');
        const dashVal = dashFilter.value;
        dashFilter.innerHTML = '<option value="All">All Teams</option>';
        teamArr.forEach(t => dashFilter.innerHTML += `<option value="${t}">${t}</option>`);
        if(teamArr.includes(dashVal)) dashFilter.value = dashVal;
        
        const setFilter = document.getElementById('settingsTeamSelect');
        const setVal = setFilter.value;
        setFilter.innerHTML = '<option value="">-- Choose a Team --</option>';
        teamArr.forEach(t => setFilter.innerHTML += `<option value="${t}">${t}</option>`);
        if(teamArr.includes(setVal)) setFilter.value = setVal;
    }

    function updateTimeUI() {
        if(!globalDataCache) return;
        const frame = document.getElementById('timeFrameFilter').value;
        const selectorContainer = document.getElementById('periodSelectorContainer');
        const periodLabel = document.getElementById('periodLabel');
        const btnPrev = document.getElementById('periodPrev');
        const btnNext = document.getElementById('periodNext');
        const intervalSelect = document.getElementById('intervalFilter');
        
        const oldInterval = intervalSelect.value;
        
        if (frame === 'All Time' || frame === 'All') {
            selectorContainer.style.display = 'none';
            intervalSelect.innerHTML = `
                <option value="Year">Yearly</option>
                <option value="Quarter">Quarterly</option>
                <option value="Month">Monthly</option>
                <option value="Week" selected>Weekly</option>
            `;
        } else {
            selectorContainer.style.display = 'inline-flex';
            let list = availablePeriods[frame];
            let idx = currentPeriodIdx[frame];
            if (!list || list.length === 0) {
                periodLabel.textContent = "N/A";
                btnPrev.disabled = true; btnNext.disabled = true;
            } else {
                periodLabel.textContent = list[idx];
                btnPrev.disabled = (idx === 0);
                btnNext.disabled = (idx === list.length - 1);
            }
            
            let opts = [];
            if (frame === 'Year') opts = ['Quarter|Quarterly', 'Month|Monthly', 'Week|Weekly'];
            if (frame === 'Quarter') opts = ['Month|Monthly', 'Week|Weekly'];
            if (frame === 'Month') opts = ['Week|Weekly'];
            
            intervalSelect.innerHTML = opts.map(o => {
                let [v, l] = o.split('|');
                return `<option value="${v}">${l}</option>`;
            }).join('');
        }
        
        let newOpts = Array.from(intervalSelect.options).map(o => o.value);
        if (newOpts.includes(oldInterval)) intervalSelect.value = oldInterval;
        else intervalSelect.value = 'Week';
        
        renderDashboard(globalDataCache);
    }

    document.getElementById('timeFrameFilter').addEventListener('change', () => { updateTimeUI(); });
    document.getElementById('intervalFilter').addEventListener('change', () => { renderDashboard(globalDataCache); });
    document.getElementById('teamFilter').addEventListener('change', () => { renderDashboard(globalDataCache); });
    document.getElementById('groupByFilter').addEventListener('change', () => { renderDashboard(globalDataCache); });

    document.getElementById('periodPrev').addEventListener('click', () => {
        let frame = document.getElementById('timeFrameFilter').value;
        if (currentPeriodIdx[frame] > 0) {
            currentPeriodIdx[frame]--;
            updateTimeUI();
        }
    });
    document.getElementById('periodNext').addEventListener('click', () => {
        let frame = document.getElementById('timeFrameFilter').value;
        if (currentPeriodIdx[frame] < availablePeriods[frame].length - 1) {
            currentPeriodIdx[frame]++;
            updateTimeUI();
        }
    });


    // --- UPLOAD VIEW ---
    let weeklyFiles = [];
    let buFile = null;

    const dtReports = document.getElementById('dropzone-reports');
    const dtBu = document.getElementById('dropzone-bu');
    const inputReports = document.getElementById('weeklyReportFile');
    const inputBu = document.getElementById('buMappingFile');
    const listReports = document.getElementById('reportsFileList');
    const listBu = document.getElementById('buFileList');

    function setupDropzone(dropzone, inputElement, isMultiple, callback) {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropzone.addEventListener(eventName, preventDefaults, false);
        });
        function preventDefaults(e) { e.preventDefault(); e.stopPropagation(); }
        ['dragenter', 'dragover'].forEach(eventName => {
            dropzone.addEventListener(eventName, () => dropzone.classList.add('dragover'), false);
        });
        ['dragleave', 'drop'].forEach(eventName => {
            dropzone.addEventListener(eventName, () => dropzone.classList.remove('dragover'), false);
        });
        dropzone.addEventListener('drop', (e) => { callback(e.dataTransfer.files); });
        inputElement.addEventListener('change', function() { callback(this.files); });
    }

    setupDropzone(dtReports, inputReports, true, (files) => {
        for(let i=0; i<files.length; i++) {
            if(files[i].name.endsWith('.xlsx') || files[i].name.endsWith('.xlsm')) {
                if(!weeklyFiles.find(f => f.name === files[i].name)) {
                    weeklyFiles.push(files[i]);
                }
            }
        }
        renderReportsList();
    });

    setupDropzone(dtBu, inputBu, false, (files) => {
        if(files.length > 0 && (files[0].name.endsWith('.xlsx') || files[0].name.endsWith('.xlsm'))) {
            buFile = files[0];
            renderBuList();
        }
    });

    function renderReportsList() {
        listReports.innerHTML = weeklyFiles.map((f, i) => `<div style="display:flex; justify-content:space-between; padding: 5px 0;"><span>📄 ${f.name} <small style="color:#666;">(${new Date(f.lastModified).toLocaleString()})</small></span><span style="color:red;cursor:pointer;font-weight:bold;" onclick="window.removeReport(${i})">X</span></div>`).join('');
    }
    window.removeReport = function(i) { weeklyFiles.splice(i, 1); renderReportsList(); };

    function renderBuList() {
        listBu.innerHTML = buFile ? `<div style="display:flex; justify-content:space-between; padding: 5px 0;"><span>📄 ${buFile.name}</span> <span style="color:red;cursor:pointer;font-weight:bold;" onclick="window.removeBu()">X</span></div>` : '';
    }
    window.removeBu = function() { buFile = null; renderBuList(); };

    document.getElementById('uploadReportsBtn').addEventListener('click', () => {
        if(weeklyFiles.length === 0) return alert("Please select weekly reports first.");
        const btn = document.getElementById('uploadReportsBtn');
        const status = document.getElementById('uploadReportsStatus');
        btn.textContent = "Uploading..."; btn.disabled = true; status.textContent = "";

        const fd = new FormData();
        weeklyFiles.sort((a, b) => a.lastModified - b.lastModified);
        weeklyFiles.forEach(f => fd.append('weekly_reports', f));

        fetch('/api/upload_reports', { method: 'POST', body: fd })
        .then(res => res.json())
        .then(data => {
            status.style.color = (data.status === 'success' || data.status === 'partial_error') ? "green" : "red";
            status.textContent = "Upload complete! Processed " + data.results.length + " files.";
            weeklyFiles = []; renderReportsList();
            loadDataAndRender();
            setTimeout(() => status.textContent = '', 4000);
        })
        .finally(() => { btn.textContent = "Upload Weekly Reports"; btn.disabled = false; });
    });

    document.getElementById('uploadBuBtn').addEventListener('click', () => {
        if(!buFile) return alert("Please select a BU Mapping file first.");
        const btn = document.getElementById('uploadBuBtn');
        const status = document.getElementById('uploadBuStatus');
        btn.textContent = "Uploading..."; btn.disabled = true; status.textContent = "";

        const fd = new FormData();
        fd.append('bu_mapping', buFile);

        fetch('/api/upload_bu', { method: 'POST', body: fd })
        .then(res => res.json())
        .then(data => {
            status.style.color = data.status === 'success' ? "green" : "red";
            status.textContent = data.message || data.error;
            if(data.status === 'success') { 
                buFile = null; 
                renderBuList(); 
                loadDataAndRender();
            }
            setTimeout(() => status.textContent = '', 4000);
        })
        .finally(() => { btn.textContent = "Upload BU Classification"; btn.disabled = false; });
    });

    document.getElementById('saveSettingsBtn').addEventListener('click', () => {
        const status = document.getElementById('settingsStatus');
        status.textContent = "Settings Saved!";
        status.style.color = "green";
        setTimeout(() => status.textContent = '', 4000);
    });

    // --- SETTINGS VIEW ---
    document.getElementById('settingsTeamSelect').addEventListener('change', (e) => {
        const team = e.target.value;
        const container = document.getElementById('teamLogContainer');
        if(!team) { container.style.display = 'none'; return; }
        container.style.display = 'block';
        renderTeamLog(team);
    });

    function renderTeamLog(team) {
        const logsData = globalDataCache.Settings.headcount_logs || [];
        const teamLog = logsData.find(l => l.department === team) || { initial: 0, logs: [] };
        
        document.getElementById('initialHeadcount').value = teamLog.initial;
        
        const tbody = document.getElementById('logTbody');
        tbody.innerHTML = '';
        teamLog.logs.forEach((log, idx) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${log.date}</td>
                <td>${log.type}</td>
                <td>${log.name}</td>
                <td><button class="btn btn-secondary remove-log" style="padding:4px 8px; font-size:0.8rem;" data-idx="${idx}">X</button></td>
            `;
            tbody.appendChild(tr);
        });
        
        document.querySelectorAll('.remove-log').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = e.target.getAttribute('data-idx');
                teamLog.logs.splice(idx, 1);
                saveSettingsRequest({ headcount_logs: updateLogsInCache(team, teamLog) });
            });
        });
    }

    document.getElementById('initialHeadcount').addEventListener('change', (e) => {
        const team = document.getElementById('settingsTeamSelect').value;
        if(!team) return;
        const logsData = globalDataCache.Settings.headcount_logs || [];
        let teamLog = logsData.find(l => l.department === team);
        if(!teamLog) {
            teamLog = { department: team, initial: parseInt(e.target.value) || 0, logs: [] };
        } else {
            teamLog.initial = parseInt(e.target.value) || 0;
        }
        saveSettingsRequest({ headcount_logs: updateLogsInCache(team, teamLog) });
    });

    document.getElementById('addLogBtn').addEventListener('click', () => {
        const team = document.getElementById('settingsTeamSelect').value;
        if(!team) return;
        
        const date = document.getElementById('newLogDate').value;
        const typeSelect = document.getElementById('newLogType');
        const typeStr = typeSelect.options[typeSelect.selectedIndex].text;
        
        let delta = 1; 
        if(typeSelect.value.includes('Leave') || typeSelect.value.includes('Depart') || typeSelect.value.includes('NoTracking')) delta = -1;
        
        const name = document.getElementById('newLogName').value;
        if(!date) return alert("Please select a date.");
        
        let teamLog = (globalDataCache.Settings.headcount_logs || []).find(l => l.department === team);
        if(!teamLog) teamLog = { department: team, initial: 0, logs: [] };
        
        teamLog.logs.push({ date, type: typeStr, delta, name });
        teamLog.logs.sort((a,b) => new Date(a.date) - new Date(b.date));
        
        saveSettingsRequest({ headcount_logs: updateLogsInCache(team, teamLog) });
        
        document.getElementById('newLogDate').value = '';
        document.getElementById('newLogName').value = '';
    });

    document.getElementById('logUploadFile').addEventListener('change', function() {
        if(this.files.length === 0) return;
        const targetTeam = document.getElementById('settingsTeamSelect').value;
        const fn = this.files[0];
        const fd = new FormData();
        fd.append('log_file', fn);
        fd.append('target_team', targetTeam || '');
        const st = document.getElementById('logUploadStatus');
        st.textContent = "Uploading & parsing..."; st.style.color = "blue";
        
        fetch('/api/upload_logs', { method: 'POST', body: fd })
        .then(res => res.json())
        .then(data => {
            if(data.status === 'success') {
                st.textContent = data.message;
                st.style.color = "green";
                loadDataAndRender();
            } else {
                st.textContent = "Error: " + data.error;
                st.style.color = "red";
            }
            setTimeout(() => st.textContent = '', 5000);
        })
        .finally(() => { document.getElementById('logUploadFile').value = ''; });
    });

    function updateLogsInCache(team, updatedTeamLog) {
        let logsData = globalDataCache.Settings.headcount_logs || [];
        const idx = logsData.findIndex(l => l.department === team);
        if(idx >= 0) logsData[idx] = updatedTeamLog;
        else logsData.push(updatedTeamLog);
        globalDataCache.Settings.headcount_logs = logsData;
        return logsData;
    }

    function renderHolidays() {
        const hList = globalDataCache.Settings.holidays || [];
        const tbody = document.getElementById('holidayTbody');
        tbody.innerHTML = '';
        hList.forEach((h, idx) => {
            const hObj = typeof h === 'string' ? {date: h, name: 'Holiday'} : h;
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${hObj.date}</td>
                <td>${hObj.name || '-'}</td>
                <td><button class="btn btn-secondary remove-holiday" style="padding:4px 8px; font-size:0.8rem;" data-idx="${idx}">X</button></td>
            `;
            tbody.appendChild(tr);
        });
        
        document.querySelectorAll('.remove-holiday').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = e.target.getAttribute('data-idx');
                hList.splice(idx, 1);
                saveSettingsRequest({ holidays: hList });
            });
        });
    }

    document.getElementById('addHolidayBtn').addEventListener('click', () => {
        const date = document.getElementById('newHolidayDate').value;
        const name = document.getElementById('newHolidayName').value;
        if(!date) return alert("Please select a date for the holiday.");
        
        const hList = globalDataCache.Settings.holidays || [];
        // prevent duplicate date
        let existing = hList.find(x => typeof x === 'string' ? x === date : x.date === date);
        if(!existing) {
            hList.push({ date, name: name || 'Holiday' });
            hList.sort((a,b) => new Date(a.date || a) - new Date(b.date || b));
            saveSettingsRequest({ holidays: hList });
        } else {
            alert("This holiday date already exists.");
        }
        
        document.getElementById('newHolidayDate').value = '';
        document.getElementById('newHolidayName').value = '';
    });

    document.getElementById('holidayUploadFile').addEventListener('change', function() {
        if(this.files.length === 0) return;
        const fn = this.files[0];
        const fd = new FormData();
        fd.append('holiday_file', fn);
        const st = document.getElementById('holidayUploadStatus');
        st.textContent = "Uploading & parsing..."; st.style.color = "blue";
        
        fetch('/api/upload_holidays', { method: 'POST', body: fd })
        .then(res => res.json())
        .then(data => {
            if(data.status === 'success') {
                st.textContent = data.message;
                st.style.color = "green";
                loadDataAndRender();
            } else {
                st.textContent = "Error: " + data.error;
                st.style.color = "red";
            }
            setTimeout(() => st.textContent = '', 5000);
        })
        .finally(() => { document.getElementById('holidayUploadFile').value = ''; });
    });

    function saveSettingsRequest(payload) {
        fetch('/api/settings', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(payload)
        }).then(res => {
            if(res.ok) loadDataAndRender();
        });
    }

    // --- DASHBOARD VIEW ---
    function computeDynamicHeadcount(team, weekDateStr) {
        const logsData = globalDataCache.Settings.headcount_logs || [];
        const teamLog = logsData.find(l => l.department === team);
        if(!teamLog) return 0;
        
        let currentHc = teamLog.initial || 0;
        let formattedStr = weekDateStr.replace(/\//g, '-');
        const targetDate = new Date(formattedStr);
        if (isNaN(targetDate)) return currentHc;

        teamLog.logs.forEach(log => {
            if(new Date(log.date) <= targetDate) { currentHc += log.delta; }
        });
        return Math.max(0, currentHc);
    }

    function renderDashboard(data) {
        if (!data || !data.WeeklyData || data.WeeklyData.length === 0) {
            document.getElementById('total-hours').textContent = '0';
            document.getElementById('overall-util').textContent = '--%';
            return;
        }
        
        const filterTeam = document.getElementById('teamFilter').value;
        const frame = document.getElementById('timeFrameFilter').value;
        const interval = document.getElementById('intervalFilter').value;
        const groupBy = document.getElementById('groupByFilter').value; 

        // Apply Dynamic Titles
        const displayTeamStr = filterTeam === 'All' ? '[All Teams]' : `[${filterTeam}]`;
        const periodStr = (frame === 'All Time' || frame === 'All') ? 'All Time' : document.getElementById('periodLabel').textContent;
        const prefix = `${displayTeamStr} [${periodStr}]`;
        const groupWord = groupBy === 'BU' ? 'BU' : 'Product Code';

        document.getElementById('buChartTitle').textContent = `${prefix} Hours Array by ${groupWord}`;
        document.getElementById('utilChartTitle').textContent = `${prefix} Workforce Utilization Rate (%)`;
        document.getElementById('pieTitle').textContent = `${prefix} Total Hours & Avg Utilization`;
        
        let weeks = [...data.WeeklyData].sort((a,b) => a.week_id.localeCompare(b.week_id));
        document.getElementById('last-updated').textContent = "Last Updated: " + weeks[weeks.length-1].date;

        if (frame !== 'All Time' && frame !== 'All') {
            const selectedPeriodLabel = availablePeriods[frame][currentPeriodIdx[frame]];
            weeks = weeks.filter(wk => {
                let mLabel = getAssignedMonthLabel(wk.date);
                if (frame === 'Month' && mLabel === selectedPeriodLabel) return true;
                if (frame === 'Quarter' && getAssignedQuarterLabel(wk.date) === selectedPeriodLabel) return true;
                if (frame === 'Year' && mLabel.split('-')[0] === selectedPeriodLabel) return true;
                return false;
            });
        }

        if (weeks.length === 0) {
            document.getElementById('total-hours').textContent = '0';
            document.getElementById('overall-util').textContent = '--%';
            if (buChartInstance) buChartInstance.destroy();
            if (utilChartInstance) utilChartInstance.destroy();
            if (pieChartInstance) pieChartInstance.destroy();
            document.getElementById('pieLegend').innerHTML = '';
            return;
        }

        const grouped = {};
        weeks.forEach(wk => {
            let label = wk.week_id;
            if (interval === 'Month') label = getAssignedMonthLabel(wk.date);
            else if (interval === 'Quarter') label = getAssignedQuarterLabel(wk.date);
            else if (interval === 'Year') label = getAssignedMonthLabel(wk.date).split('-')[0];
            
            if(!grouped[label]) grouped[label] = { weeks: [] };
            grouped[label].weeks.push(wk);
        });

        // Sort chronologically based on the first recorded week's pure date inside the grouping
        const groupedTokens = Object.keys(grouped).map(lbl => {
            return {
                lbl: lbl,
                firstDate: new Date(grouped[lbl].weeks[0].date.replace(/\//g, '-'))
            };
        });
        groupedTokens.sort((a,b) => a.firstDate - b.firstDate);
        const labels = groupedTokens.map(g => g.lbl);
        
        let totalHoursAllTime = 0;
        let grandUtilSum = 0;
        let grandUtilMax = 0;

        const buTrends = {}; 
        const utilTrends = {}; 
        const pieData = {};
        const pcToBu = {}; // Used when groupBy === 'Product Code' to maintain BU linkage

        const holidays = data.Settings.holidays || [];

        labels.forEach((lbl, LIdx) => {
            let labelData = grouped[lbl].weeks;
            let periodBuHours = {};
            let periodDeptLogged = {};
            let periodDeptMax = {};

            labelData.forEach(wk => {
                let workDays = wk.working_days || 5;
                let weekStart = new Date(wk.date.replace(/\//g, '-'));
                let weekEnd = new Date(weekStart);
                weekEnd.setDate(weekStart.getDate() + 4); 
                let holidayCount = 0;
                holidays.forEach(h => {
                    const hDate = new Date(typeof h === 'string' ? h : h.date);
                    if (hDate >= weekStart && hDate <= weekEnd) holidayCount++;
                });
                workDays = Math.max(0, workDays - holidayCount);

                let wkDeptLogged = {};
                
                wk.records.forEach(r => {
                    if(filterTeam !== 'All' && r.department !== filterTeam) return;
                    
                    let groupKey = groupBy === 'BU' ? (r.bu || "Unknown") : (r.product_code || "Unknown");
                    const dpt = r.department;
                    
                    if(!periodBuHours[groupKey]) periodBuHours[groupKey] = 0;
                    periodBuHours[groupKey] += r.hours;
                    
                    if(!wkDeptLogged[dpt]) wkDeptLogged[dpt] = 0;
                    wkDeptLogged[dpt] += r.hours;
                    
                    totalHoursAllTime += r.hours;
                    if(!pieData[groupKey]) pieData[groupKey] = 0;
                    pieData[groupKey] += r.hours;
                    
                    if(groupBy === 'Product Code' && !pcToBu[groupKey]) {
                        pcToBu[groupKey] = r.bu || 'Unknown';
                    }
                });
                
                Object.keys(wkDeptLogged).forEach(dpt => {
                    const hc = computeDynamicHeadcount(dpt, wk.date);
                    if (hc > 0) {
                        const wkMax = hc * workDays * 8;
                        if(!periodDeptLogged[dpt]) periodDeptLogged[dpt] = 0;
                        if(!periodDeptMax[dpt]) periodDeptMax[dpt] = 0;
                        periodDeptLogged[dpt] += wkDeptLogged[dpt];
                        periodDeptMax[dpt] += wkMax;
                        
                        grandUtilSum += wkDeptLogged[dpt];
                        grandUtilMax += wkMax;
                    }
                });
            });

            Object.keys(periodBuHours).forEach(k => {
                 if(!buTrends[k]) buTrends[k] = new Array(labels.length).fill(0);
                 buTrends[k][LIdx] = periodBuHours[k];
            });
            
            Object.keys(periodDeptMax).forEach(dpt => {
                 if(!utilTrends[dpt]) utilTrends[dpt] = new Array(labels.length).fill(null);
                 utilTrends[dpt][LIdx] = periodDeptMax[dpt] > 0 ? (periodDeptLogged[dpt] / periodDeptMax[dpt]) * 100 : 0;
            });
        });

        document.getElementById('total-hours').textContent = (Math.round(totalHoursAllTime * 10) / 10).toFixed(1);
        document.getElementById('overall-util').textContent = grandUtilMax > 0 ? ((grandUtilSum/grandUtilMax)*100).toFixed(1) + "%" : "--%";

        const colorPalette = [
            'rgba(67, 97, 238, 0.5)', 'rgba(76, 201, 240, 0.5)', 'rgba(247, 37, 133, 0.5)', 
            'rgba(114, 9, 183, 0.5)', 'rgba(58, 12, 163, 0.5)', 'rgba(252, 191, 73, 0.5)',
            'rgba(46, 196, 182, 0.5)', 'rgba(231, 111, 81, 0.5)', 'rgba(233, 196, 106, 0.5)'
        ];
        const borderPalette = [
            '#4361ee', '#4cc9f0', '#f72585', '#7209b7', '#3a0ca3', '#fcbf49',
            '#2ec4b6', '#e76f51', '#e9c46a'
        ];

        let colorIdx = 0;
        const buDatasets = Object.keys(buTrends).sort().map(k => {
            const bgC = colorPalette[colorIdx % colorPalette.length];
            const bdC = borderPalette[colorIdx % borderPalette.length];
            colorIdx++;
            return {
                label: k,
                data: buTrends[k],
                fill: true,
                backgroundColor: bgC,
                borderColor: bdC,
                borderWidth: 2,
                tension: 0.3
            };
        });

        if (buChartInstance) buChartInstance.destroy();
        const buCtx = document.getElementById('buChart').getContext('2d');
        buChartInstance = new Chart(buCtx, {
            type: 'line',
            data: { labels: labels, datasets: buDatasets },
            options: { 
                responsive: true, 
                maintainAspectRatio: false,
                plugins: { title: { display: false } },
                scales: {
                    x: { title: { display: true, text: interval } },
                    y: { stacked: true, beginAtZero: true, title: { display: true, text: 'Hours' } }
                }
            }
        });

        const utilDatasets = Object.keys(utilTrends).sort().map((dpt, idx) => ({
            label: dpt,
            data: utilTrends[dpt],
            fill: false,
            borderColor: borderPalette[idx % borderPalette.length],
            tension: 0.3,
            borderWidth: 3
        }));

        if (utilChartInstance) utilChartInstance.destroy();
        const utilCtx = document.getElementById('utilChart').getContext('2d');
        utilChartInstance = new Chart(utilCtx, {
            type: 'line',
            data: { labels: labels, datasets: utilDatasets },
            options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true } } }
        });

        // Group Pie Chart
        let pieSortedKeys = Object.keys(pieData).sort();
        let pieColors = pieSortedKeys.map((_, i) => borderPalette[i % borderPalette.length]);
        let pieBgColors = pieSortedKeys.map((_, i) => colorPalette[i % colorPalette.length]);
        
        if (pieChartInstance) pieChartInstance.destroy();
        const pieCtx = document.getElementById('buPieChart').getContext('2d');
        pieChartInstance = new Chart(pieCtx, {
            type: 'doughnut',
            data: {
                labels: pieSortedKeys,
                datasets: [{
                    data: pieSortedKeys.map(k => pieData[k]),
                    backgroundColor: pieBgColors,
                    borderColor: '#ffffff',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                cutout: '60%',
                plugins: { legend: { display: false } }
            }
        });
        
        const legendContainer = document.getElementById('pieLegend');
        legendContainer.innerHTML = '';
        
        if (groupBy === 'BU') {
            legendContainer.style.display = 'flex';
            legendContainer.style.flexDirection = 'column';
            legendContainer.style.flexWrap = 'wrap';
            legendContainer.style.overflowX = 'hidden';
            legendContainer.style.alignContent = 'flex-start';
            legendContainer.style.maxHeight = '300px';
            legendContainer.style.gap = '8px 25px';

            pieSortedKeys.forEach((k, i) => {
                let hrsRaw = pieData[k];
                let hrsRounded = (Math.round(hrsRaw * 10) / 10).toFixed(1);
                let pct = totalHoursAllTime > 0 ? ((hrsRaw / totalHoursAllTime) * 100).toFixed(1) : 0;
                legendContainer.innerHTML += `
                    <div style="display:flex; align-items:center; gap:8px;">
                        <div style="width:14px; height:14px; background:${pieColors[i % pieColors.length]}; border-radius:3px; flex-shrink: 0;"></div>
                        <span style="font-weight:600; font-size:0.85rem; min-width:50px;">${k}</span>
                        <span style="color:#666; font-size:0.85rem;">${pct}% (${hrsRounded}h)</span>
                    </div>
                `;
            });
        } else {
            // Group by Product Code: Nested Legend
            legendContainer.style.display = 'flex';
            legendContainer.style.flexDirection = 'row';
            legendContainer.style.flexWrap = 'nowrap';
            legendContainer.style.overflowX = 'auto'; // Allow horizontal scrolling when expanded
            legendContainer.style.alignItems = 'flex-start';
            legendContainer.style.gap = '25px';

            let groupedByBu = {};
            pieSortedKeys.forEach((pc, i) => {
                let bu = pcToBu[pc] || 'Unknown';
                if(!groupedByBu[bu]) groupedByBu[bu] = [];
                groupedByBu[bu].push({ pc, idx: i });
            });
            Object.keys(groupedByBu).sort().forEach(bu => {
                let colHtml = `<div style="display:flex; flex-direction:column; min-width:130px;">`;
                colHtml += `<div style="font-weight:600; font-size: 0.95rem; color: #2b2b36; margin-top:8px; border-bottom:1px solid #e5e7eb; padding-bottom:4px; margin-bottom:8px;">${bu}</div>`;
                groupedByBu[bu].forEach(item => {
                    let hrsRaw = pieData[item.pc];
                    let hrsRounded = (Math.round(hrsRaw * 10) / 10).toFixed(1);
                    let pct = totalHoursAllTime > 0 ? ((hrsRaw / totalHoursAllTime) * 100).toFixed(1) : 0;
                    colHtml += `
                        <div style="display:flex; align-items:flex-start; gap:6px; margin-bottom:10px;">
                            <div style="width:12px; height:12px; margin-top: 2px; flex-shrink:0; background:${pieBgColors[item.idx]}; border:1px solid ${pieColors[item.idx]}; border-radius:2px;"></div>
                            <span style="font-weight:500; font-size:0.85rem; line-height:1.3;">
                                ${item.pc}
                                <span style="display:block; color:#666; font-weight:normal; margin-top:3px;">${pct}% (${hrsRounded}h)</span>
                            </span>
                        </div>
                    `;
                });
                colHtml += `</div>`;
                legendContainer.innerHTML += colHtml;
            });
        }
    }

    // Initialize application data
    loadDataAndRender();
});
