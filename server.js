/**
 * PowerPoint AI Presenter - Complete Node.js Backend Server
 * High performance, zero external npm dependencies, native Windows PowerPoint control.
 * Supports opening PPTs by name, number, ordinal (1st, 2nd, 3rd, last), or click.
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

const PORT = 3000;
const PRESENTATIONS_DIR = path.resolve('./presentations');

if (!fs.existsSync(PRESENTATIONS_DIR)) {
    fs.mkdirSync(PRESENTATIONS_DIR, { recursive: true });
}

let activePresentation = null;
let currentSlide = 1;
let totalSlides = 5;

const ORDINAL_MAP = {
    'first': 1, '1st': 1, 'one': 1, '1': 1,
    'second': 2, '2nd': 2, 'two': 2, '2': 2,
    'third': 3, '3rd': 3, 'three': 3, '3': 3,
    'fourth': 4, '4th': 4, 'four': 4, '4': 4,
    'fifth': 5, '5th': 5, 'five': 5, '5': 5,
    'sixth': 6, '6th': 6, 'six': 6, '6': 6,
    'seventh': 7, '7th': 7, 'seven': 7, '7': 7,
    'eighth': 8, '8th': 8, 'eight': 8, '8': 8,
    'ninth': 9, '9th': 9, 'nine': 9, '9': 9,
    'tenth': 10, '10th': 10, 'ten': 10, '10': 10,
    'last': 'last'
};

/**
 * List all presentation files in folder
 */
function listPresentations() {
    if (!fs.existsSync(PRESENTATIONS_DIR)) return [];
    const files = fs.readdirSync(PRESENTATIONS_DIR);
    return files
        .filter(f => f.toLowerCase().endsWith('.pptx') || f.toLowerCase().endsWith('.ppt') || f.toLowerCase().endsWith('.ppsx'))
        .filter(f => !f.startsWith('~$'))
        .sort((a, b) => a.localeCompare(b))
        .map((f, index) => {
            const fullPath = path.join(PRESENTATIONS_DIR, f);
            const stats = fs.statSync(fullPath);
            return {
                index: index + 1,
                name: f,
                path: fullPath,
                size_kb: (stats.size / 1024).toFixed(1),
                is_active: activePresentation === fullPath
            };
        });
}

/**
 * Send Windows keystrokes to PowerPoint using PowerShell WScript.Shell
 */
function sendPowerPointKey(key, callback) {
    const psScript = `
    $wshell = New-Object -ComObject WScript.Shell;
    $activated = $wshell.AppActivate('PowerPoint');
    if (-not $activated) {
        $activated = $wshell.AppActivate('Slide Show');
    }
    Start-Sleep -Milliseconds 80;
    $wshell.SendKeys('${key}');
    `;
    
    exec(`powershell -NoProfile -Command "${psScript.replace(/\r?\n/g, ' ')}"`, (err) => {
        if (err) {
            console.error('Key error:', err);
            if (callback) callback(false, err);
        } else {
            console.log(`[PowerPoint Key] '${key}' sent successfully.`);
            if (callback) callback(true);
        }
    });
}

/**
 * Launch PowerPoint presentation by Index, Name, or Ordinal
 */
function launchPowerPoint(targetNameOrIndex, callback) {
    const files = listPresentations();
    if (files.length === 0) {
        if (callback) callback(false, "No PowerPoint files found in presentations/ folder.");
        return;
    }

    let targetFile = null;

    // 1. If target is a number (e.g. 2, 3)
    if (typeof targetNameOrIndex === 'number') {
        const idx = Math.max(1, Math.min(targetNameOrIndex, files.length)) - 1;
        targetFile = files[idx].path;
    } else if (targetNameOrIndex) {
        const clean = targetNameOrIndex.toString().trim().toLowerCase();

        // Check if ordinal ("second", "2nd", "3rd", "last")
        if (clean === 'last') {
            targetFile = files[files.length - 1].path;
        } else if (ORDINAL_MAP[clean]) {
            const idxNum = ORDINAL_MAP[clean];
            const idx = Math.max(1, Math.min(idxNum, files.length)) - 1;
            targetFile = files[idx].path;
        } else if (!['', 'ppt', 'the ppt', 'presentation', 'the presentation', 'deck'].includes(clean)) {
            // Check direct number inside text (e.g., "ppt 2", "2")
            const digitMatch = clean.match(/\b(\d+)\b/);
            if (digitMatch) {
                const num = parseInt(digitMatch[1], 10);
                const idx = Math.max(1, Math.min(num, files.length)) - 1;
                targetFile = files[idx].path;
            } else {
                // Fuzzy search by filename substring
                const found = files.find(f => f.name.toLowerCase().includes(clean));
                if (found) {
                    targetFile = found.path;
                }
            }
        }
    }

    // Default to first file if no target matched
    if (!targetFile) {
        targetFile = files[0].path;
    }

    activePresentation = targetFile;
    currentSlide = 1;
    console.log(`[Launch PPT] Opening: ${targetFile}`);

    // Launch with Windows start command
    exec(`cmd /c start "" "${targetFile}"`, (err) => {
        if (err) {
            console.error('Launch error:', err);
            exec(`powershell -Command "Start-Process '${targetFile}'"`, (psErr) => {
                if (psErr && callback) callback(false, psErr);
                else if (callback) callback(true, path.basename(targetFile));
            });
        } else {
            console.log(`[Launch PPT] PowerPoint launched.`);
            setTimeout(() => {
                sendPowerPointKey('{F5}');
            }, 1500);
            if (callback) callback(true, path.basename(targetFile));
        }
    });
}

/**
 * Natural Language Command Parser
 */
function parseCommand(rawText) {
    if (!rawText || !rawText.trim()) {
        return { action: 'unknown', feedback: "I didn't catch that. Please speak or type a command." };
    }

    const text = rawText.toLowerCase().replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim();

    // 0. List Presentations
    if (/\b(list|show|what|which)\b/.test(text) && /\b(presentations|ppts|decks|files)\b/.test(text)) {
        const files = listPresentations();
        const listStr = files.map(f => `#${f.index}: ${f.name}`).join(', ');
        return {
            action: 'list_decks',
            feedback: `Available presentations: ${listStr}. Say "Open second PPT" or "Open [name]" to launch.`
        };
    }

    // 1. Open PPT (e.g. "open 2nd ppt", "open second ppt", "open 3rd presentation", "open sample", "open ppt number 3")
    if (/\b(open|launch|start|load|play|show)\b/.test(text)) {
        // Check for specific ordinal / index patterns:
        // "open second ppt", "open 2nd ppt", "open ppt 2", "open 3rd presentation", "open third ppt"
        const ordinalMatch = text.match(/\b(first|1st|second|2nd|third|3rd|fourth|4th|fifth|5th|sixth|6th|seventh|7th|eighth|8th|ninth|9th|tenth|10th|last)\b/);
        if (ordinalMatch) {
            const ord = ordinalMatch[1];
            return {
                action: 'open',
                target: ord,
                feedback: `Opening the ${ord} presentation...`
            };
        }

        const numMatch = text.match(/\b(?:ppt|presentation|deck|number|no\.?)\s+(\d+)\b/) || text.match(/\b(?:open|load)\s+(\d+)\b/);
        if (numMatch) {
            const num = parseInt(numMatch[1], 10);
            return {
                action: 'open',
                target: num,
                feedback: `Opening presentation #${num}...`
            };
        }

        // Generic / named matching: "open sample", "open report"
        let match = text.replace(/\b(open|the|launch|start|load|ppt|presentation|deck|please|now|file)\b/g, '').trim();
        return {
            action: 'open',
            target: match || null,
            feedback: match ? `Opening presentation matching "${match}"...` : "Opening PowerPoint presentation..."
        };
    }

    // 2. Next Slide
    if (/\b(next|advance|forward|next slide|next page|move forward)\b/.test(text)) {
        return { action: 'next', feedback: "Moving to next slide." };
    }

    // 3. Previous Slide / Back
    if (/\b(previous|prev|back|go back|last slide was|previous slide)\b/.test(text) && !/\blast slide\b/.test(text)) {
        return { action: 'prev', feedback: "Going back to previous slide." };
    }

    // 4. First Slide / Beginning
    if (/\b(first slide|beginning|start from beginning|go to start)\b/.test(text)) {
        return { action: 'first', feedback: "Navigating to first slide." };
    }

    // 5. Last Slide / End
    if (/\b(last slide|end of presentation|final slide|go to end)\b/.test(text)) {
        return { action: 'last', feedback: "Navigating to final slide." };
    }

    // 6. Jump to Slide Number
    const numMatch = text.match(/\bslide\s+(\d+)\b/) || text.match(/\b(?:go to|jump to)\s+slide\s+(\d+)\b/);
    if (numMatch) {
        const slideNum = parseInt(numMatch[1], 10);
        return { action: 'goto', slide: slideNum, feedback: `Jumping to slide ${slideNum}.` };
    }

    // 7. Fullscreen / Start presentation
    if (/\b(fullscreen|full screen|present|start presentation|slideshow)\b/.test(text)) {
        return { action: 'start_show', feedback: "Starting fullscreen presentation." };
    }

    // 8. Exit / Close
    if (/\b(exit|stop|close|quit)\b/.test(text)) {
        return { action: 'stop_show', feedback: "Exiting presentation." };
    }

    // 9. Blackout / Blank Screen
    if (/\b(black|blackout|black screen|blank)\b/.test(text)) {
        return { action: 'blank', feedback: "Blacking out screen." };
    }

    return {
        action: 'unknown',
        feedback: `Command "${rawText}" not recognized. Try saying "Open 2nd PPT", "Next slide", or "List presentations".`
    };
}

/**
 * Execute Action
 */
function handleAction(action, params, res) {
    if (action === 'open') {
        launchPowerPoint(params ? params.target : null, (ok, detail) => {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({
                success: ok,
                feedback: ok ? `Opened presentation: ${detail}` : `Error: ${detail}`,
                activeDeck: activePresentation ? path.basename(activePresentation) : null,
                currentSlide
            }));
        });
        return;
    }

    if (action === 'list_decks') {
        const files = listPresentations();
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
            success: true,
            feedback: params.feedback || `Found ${files.length} presentations.`,
            presentations: files
        }));
        return;
    }

    if (action === 'next') {
        currentSlide++;
        sendPowerPointKey('{RIGHT}', (ok) => {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({
                success: ok,
                feedback: `Slide ${currentSlide}`,
                currentSlide
            }));
        });
        return;
    }

    if (action === 'prev') {
        if (currentSlide > 1) currentSlide--;
        sendPowerPointKey('{LEFT}', (ok) => {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({
                success: ok,
                feedback: `Slide ${currentSlide}`,
                currentSlide
            }));
        });
        return;
    }

    if (action === 'first') {
        currentSlide = 1;
        sendPowerPointKey('{HOME}', (ok) => {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: ok, feedback: "At the first slide.", currentSlide }));
        });
        return;
    }

    if (action === 'last') {
        sendPowerPointKey('{END}', (ok) => {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: ok, feedback: "At the last slide.", currentSlide }));
        });
        return;
    }

    if (action === 'goto') {
        const slide = (params && params.slide) ? params.slide : 1;
        currentSlide = slide;
        sendPowerPointKey(`${slide}{ENTER}`, (ok) => {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: ok, feedback: `Jumped to slide ${slide}.`, currentSlide }));
        });
        return;
    }

    if (action === 'start_show') {
        sendPowerPointKey('{F5}', (ok) => {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: ok, feedback: "Started presentation.", currentSlide }));
        });
        return;
    }

    if (action === 'stop_show') {
        sendPowerPointKey('{ESC}', (ok) => {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: ok, feedback: "Exited slideshow.", currentSlide }));
        });
        return;
    }

    if (action === 'blank') {
        sendPowerPointKey('b', (ok) => {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: ok, feedback: "Toggled blackout screen.", currentSlide }));
        });
        return;
    }

    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ success: false, feedback: `Unknown action ${action}` }));
}

/**
 * HTTP Server
 */
const server = http.createServer((req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }

    const parsedUrl = new URL(req.url, `http://${req.headers.host}`);
    const pathname = parsedUrl.pathname;

    // 1. API: List Presentations
    if (pathname === '/api/presentations' && req.method === 'GET') {
        const files = listPresentations();
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
            folder: PRESENTATIONS_DIR,
            count: files.length,
            presentations: files
        }));
        return;
    }

    // 2. API: Status
    if (pathname === '/api/status' && req.method === 'GET') {
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
            running: !!activePresentation,
            presentation_name: activePresentation ? path.basename(activePresentation) : "No presentation active",
            current_slide: currentSlide,
            total_slides: totalSlides,
            folder: PRESENTATIONS_DIR
        }));
        return;
    }

    // 3. API: Natural Language Command
    if (pathname === '/api/command' && req.method === 'POST') {
        let body = '';
        req.on('data', chunk => { body += chunk; });
        req.on('end', () => {
            try {
                const data = JSON.parse(body);
                const parsed = parseCommand(data.text || '');
                console.log(`[Command Received] "${data.text}" -> Action: ${parsed.action} (Target: ${parsed.target || 'none'})`);

                if (parsed.action === 'unknown') {
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({
                        success: false,
                        feedback: parsed.feedback,
                        currentSlide
                    }));
                    return;
                }

                handleAction(parsed.action, { target: parsed.target, slide: parsed.slide, feedback: parsed.feedback }, res);
            } catch (err) {
                res.writeHead(400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ success: false, error: err.message }));
            }
        });
        return;
    }

    // 4. API: Direct Action
    if (pathname === '/api/action' && req.method === 'POST') {
        let body = '';
        req.on('data', chunk => { body += chunk; });
        req.on('end', () => {
            try {
                const data = JSON.parse(body);
                handleAction(data.action, data.params || {}, res);
            } catch (err) {
                res.writeHead(400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ success: false, error: err.message }));
            }
        });
        return;
    }

    // 5. Static Files Serving
    let filePath = path.join(__dirname, pathname === '/' ? 'static/index.html' : pathname);
    if (pathname.startsWith('/static/')) {
        filePath = path.join(__dirname, pathname);
    }

    if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
        const ext = path.extname(filePath).toLowerCase();
        const mimeTypes = {
            '.html': 'text/html',
            '.js': 'text/javascript',
            '.css': 'text/css',
            '.json': 'application/json',
            '.svg': 'image/svg+xml'
        };
        const contentType = mimeTypes[ext] || 'application/octet-stream';
        res.writeHead(200, { 'Content-Type': contentType });
        fs.createReadStream(filePath).pipe(res);
    } else {
        const indexPath = path.join(__dirname, 'static/index.html');
        if (fs.existsSync(indexPath)) {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            fs.createReadStream(indexPath).pipe(res);
        } else {
            res.writeHead(404, { 'Content-Type': 'text/plain' });
            res.end('Not Found');
        }
    }
});

const os = require('os');

function getLocalIPAddresses() {
    const interfaces = os.networkInterfaces();
    const ips = [];
    for (const name of Object.keys(interfaces)) {
        for (const iface of interfaces[name]) {
            if (iface.family === 'IPv4' && !iface.internal) {
                ips.push(iface.address);
            }
        }
    }
    return ips;
}

server.listen(PORT, '0.0.0.0', () => {
    const localIPs = getLocalIPAddresses();
    console.log('='.repeat(65));
    console.log('  🎯 Node.js PowerPoint AI Voice & Chat Assistant');
    console.log('='.repeat(65));
    console.log(`[*] Localhost:        http://127.0.0.1:${PORT}`);
    if (localIPs.length > 0) {
        localIPs.forEach(ip => {
            console.log(`[*] Local Network IP: http://${ip}:${PORT}`);
        });
    }
    console.log(`[*] Presentations:    ${PRESENTATIONS_DIR}`);
    console.log('='.repeat(65));
});

