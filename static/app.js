// DXF to GeoPackage Converter
'use strict';

const $ = (sel) => document.querySelector(sel);
let currentJobId = null;
let eventSource = null;

const dropZone = $('#drop-zone');
const fileInput = $('#file-input');
const fileInfo = $('#file-info');
const settingsPanel = $('#settings-panel');
const progressSection = $('#progress-section');
const logSection = $('#log-section');
const downloadSection = $('#download-section');
const convertBtn = $('#convert-btn');
const resetBtn = $('#reset-btn');
const clearBtn = $('#clear-btn');
const zoneSelect = $('#zone-select');
const outputCrs = $('#output-crs');
const customEpsgRow = $('#custom-epsg-row');
const statusText = $('#status-text');

// --- Status Bar ---
function setStatus(msg) {
    statusText.textContent = msg;
}

// --- Mascot ---
function mascotSay(msg) {
    const speech = $('#mascot-speech');
    if (speech) {
        speech.textContent = msg;
        speech.style.opacity = '1';
        clearTimeout(speech._timer);
        speech._timer = setTimeout(() => { speech.style.opacity = '0'; }, 4000);
    }
}

function mascotMood(mood) {
    const m = $('#mascot');
    if (m) {
        m.className = mood || '';
        if (mood === 'happy') {
            setTimeout(() => { m.className = ''; }, 3500);
        }
    }
}

// --- Upload ---
dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) uploadFile(e.target.files[0]);
});

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
    mascotSay('Drop it!');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) uploadFile(e.dataTransfer.files[0]);
});

async function uploadFile(file) {
    if (!file.name.toLowerCase().endsWith('.dxf')) {
        mascotSay('.dxf only!');
        alert('.dxf files only');
        return;
    }
    resetUI();
    setStatus('Uploading...');
    mascotSay('Reading...');

    const formData = new FormData();
    formData.append('file', file);
    $('#drop-text').textContent = 'Uploading...';

    try {
        const res = await fetch('/api/upload', { method: 'POST', body: formData });
        const data = await res.json();
        if (!res.ok) {
            alert(data.error || 'Upload error');
            $('#drop-text').textContent = 'Drop .dxf file here or click to select';
            setStatus('Ready');
            mascotSay('Hmm...');
            return;
        }

        currentJobId = data.job_id;
        $('#info-filename').textContent = data.filename;
        $('#info-filesize').textContent = formatSize(data.filesize);

        const cr = data.coord_range;
        if (cr) {
            $('#info-coords').textContent =
                `X: ${cr.x_min?.toFixed(0)}..${cr.x_max?.toFixed(0)}  ` +
                `Y: ${cr.y_min?.toFixed(0)}..${cr.y_max?.toFixed(0)}`;
        } else {
            $('#info-coords').textContent = '-';
        }

        // Populate zone select
        if (data.zones) {
            zoneSelect.innerHTML = '';
            for (const [num, desc] of Object.entries(data.zones)) {
                const opt = document.createElement('option');
                opt.value = num;
                opt.textContent = `${num} - ${desc}`;
                zoneSelect.appendChild(opt);
            }
        }

        // Auto-detect zone
        const detected = data.detected_zone;
        if (detected) {
            zoneSelect.value = String(detected);
            $('#info-zone').textContent = `${detected} (auto)`;
            mascotSay(`Zone ${detected}!`);
        } else {
            zoneSelect.value = '9';
            $('#info-zone').textContent = '9 (default)';
            mascotSay('Loaded!');
        }

        fileInfo.classList.remove('hidden');
        settingsPanel.classList.remove('hidden');
        clearBtn.classList.remove('hidden');
        dropZone.classList.add('has-file');
        $('#drop-text').textContent = data.filename;
        setStatus(`Loaded: ${data.filename}`);

    } catch (err) {
        alert('Upload failed: ' + err.message);
        $('#drop-text').textContent = 'Drop .dxf file here or click to select';
        setStatus('Ready');
        mascotSay('Error!');
    }
}

// --- Settings ---
outputCrs.addEventListener('change', () => {
    customEpsgRow.style.display = outputCrs.value === 'custom' ? '' : 'none';
});

// --- Convert ---
convertBtn.addEventListener('click', startConvert);

async function startConvert() {
    if (!currentJobId) return;
    convertBtn.disabled = true;
    convertBtn.textContent = 'Converting...';
    setStatus('Converting...');
    mascotMood('thinking');
    mascotSay('Working...');

    progressSection.classList.remove('hidden');
    logSection.classList.remove('hidden');
    downloadSection.classList.add('hidden');
    $('#progress-bar').style.width = '0%';
    $('#progress-text').textContent = '0%';
    $('#log-output').innerHTML = '';

    const params = {
        scale: parseInt($('#scale-input').value) || 0,
        datum: $('#datum-select').value,
        zone: parseInt(zoneSelect.value),
        output_crs: outputCrs.value,
        custom_epsg: parseInt($('#custom-epsg').value) || 0,
        quality: parseInt($('#quality-select').value),
        auto_georef: $('#auto-georef').checked,
        split_by_layer: $('#split-layers').checked,
    };

    try {
        const res = await fetch(`/api/convert/${currentJobId}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(params),
        });
        if (!res.ok) {
            const data = await res.json();
            alert(data.error || 'Convert error');
            convertBtn.disabled = false;
            convertBtn.textContent = 'Convert';
            setStatus('Error');
            mascotMood('');
            mascotSay('Oops!');
            return;
        }
        connectSSE();
    } catch (err) {
        alert('Convert failed: ' + err.message);
        convertBtn.disabled = false;
        convertBtn.textContent = 'Convert';
        setStatus('Error');
        mascotMood('');
        mascotSay('Failed!');
    }
}

function connectSSE() {
    if (eventSource) eventSource.close();
    eventSource = new EventSource(`/api/progress/${currentJobId}`);

    eventSource.onmessage = (event) => {
        const msg = JSON.parse(event.data);
        if (msg.type === 'log') {
            appendLog(msg.msg, msg.level);
        } else if (msg.type === 'progress') {
            setProgress(msg.value);
            setStatus(`Converting... ${msg.value}%`);
            if (msg.value >= 90) mascotSay('Almost there!');
        } else if (msg.type === 'complete') {
            setProgress(100);
            appendLog('Done.', 'info');
            showDownloads(msg.files);
            convertBtn.disabled = false;
            convertBtn.textContent = 'Convert';
            setStatus('Complete');
            mascotMood('happy');
            mascotSay('Done!');
            eventSource.close();
        } else if (msg.type === 'error') {
            appendLog('ERROR: ' + msg.msg, 'error');
            convertBtn.disabled = false;
            convertBtn.textContent = 'Convert';
            setStatus('Error');
            mascotMood('');
            mascotSay('Error...');
            eventSource.close();
        }
    };

    eventSource.onerror = () => {
        eventSource.close();
        convertBtn.disabled = false;
        convertBtn.textContent = 'Convert';
        mascotMood('');
    };
}

// --- UI ---
function appendLog(msg, level = 'info') {
    const out = $('#log-output');
    const line = document.createElement('div');
    line.className = `log-line ${level}`;
    const ts = new Date().toLocaleTimeString('en-GB', { hour12: false });
    line.textContent = `${ts}  ${msg}`;
    out.appendChild(line);
    out.scrollTop = out.scrollHeight;
}

function setProgress(v) {
    $('#progress-bar').style.width = `${v}%`;
    $('#progress-text').textContent = `${v}%`;
}

function showDownloads(files) {
    downloadSection.classList.remove('hidden');
    const c = $('#download-links');
    c.innerHTML = '';
    for (const f of files) {
        const a = document.createElement('a');
        a.href = `/api/download/${currentJobId}/${encodeURIComponent(f)}`;
        a.className = 'download-link';
        a.textContent = `>> ${f}`;
        a.download = f;
        c.appendChild(a);
    }
}

function formatSize(b) {
    if (b < 1024) return b + ' B';
    if (b < 1048576) return (b / 1024).toFixed(1) + ' KB';
    return (b / 1048576).toFixed(1) + ' MB';
}

function resetUI() {
    if (currentJobId) fetch(`/api/cleanup/${currentJobId}`, { method: 'POST' }).catch(() => {});
    if (eventSource) { eventSource.close(); eventSource = null; }
    currentJobId = null;
    fileInfo.classList.add('hidden');
    settingsPanel.classList.add('hidden');
    progressSection.classList.add('hidden');
    logSection.classList.add('hidden');
    downloadSection.classList.add('hidden');
    clearBtn.classList.add('hidden');
    convertBtn.disabled = false;
    convertBtn.textContent = 'Convert';
    fileInput.value = '';
    dropZone.classList.remove('has-file');
    setStatus('Ready');
    mascotMood('');
}

function doClear() {
    resetUI();
    $('#drop-text').textContent = 'Drop .dxf file here or click to select';
    mascotSay('Cleared!');
}

// Clear / Reset buttons
clearBtn.addEventListener('click', doClear);
resetBtn.addEventListener('click', doClear);

// --- Menu Bar ---
$('#menu-open')?.addEventListener('click', () => fileInput.click());
$('#menu-clear')?.addEventListener('click', doClear);
$('#menu-reset')?.addEventListener('click', () => location.reload());

// About dialog
$('#menu-about')?.addEventListener('click', () => {
    const mem = performance?.memory;
    if (mem) {
        $('#about-mem').textContent =
            `Memory: ${(mem.usedJSHeapSize / 1024 / 1024).toFixed(1)} MB used`;
    }
    $('#about-overlay').classList.remove('hidden');
});
$('#about-ok')?.addEventListener('click', () => {
    $('#about-overlay').classList.add('hidden');
});
$('#about-overlay')?.addEventListener('click', (e) => {
    if (e.target === e.currentTarget) $('#about-overlay').classList.add('hidden');
});

// --- Keyboard Shortcuts ---
document.addEventListener('keydown', (e) => {
    if (e.metaKey || e.ctrlKey) {
        if (e.key === 'o' || e.key === 'O') {
            e.preventDefault();
            fileInput.click();
        }
    }
    if (e.key === 'Escape') {
        $('#about-overlay')?.classList.add('hidden');
    }
});

// --- Mascot Character (created dynamically) ---
(function initMascot() {
    const mascot = document.createElement('div');
    mascot.id = 'mascot';
    mascot.innerHTML = `
        <div class="mascot-sprite">
            <div class="invader"></div>
        </div>
        <div class="mascot-speech" id="mascot-speech">Hello!</div>
    `;
    document.body.appendChild(mascot);

    // Blink at random intervals
    setInterval(() => {
        mascot.classList.add('blink');
        setTimeout(() => mascot.classList.remove('blink'), 200);
    }, 3000 + Math.random() * 2000);

    // Click to get random message
    mascot.addEventListener('click', () => {
        const msgs = [
            'Hello!', 'Drop a DXF!', 'Nice day!',
            'GeoPackage!', 'JPC Zone 9!', 'Affine!',
            'Let\'s convert!', 'System 7!',
        ];
        mascotSay(msgs[Math.floor(Math.random() * msgs.length)]);
    });

    // Hide speech after initial delay
    setTimeout(() => {
        const speech = $('#mascot-speech');
        if (speech) speech.style.opacity = '0';
    }, 3000);
})();
