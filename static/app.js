/**
 * PowerPoint AI Presenter - Voice & Chat Controller + Live Presentation Stage
 */

// Application State
const state = {
    listening: false,
    continuous: true,
    ttsEnabled: true,
    activeDeck: null,
    currentSlide: 1,
    totalSlides: 1,
    slides: [],
    blackout: false,
    folder: "./presentations",
    presentations: [],
    recognition: null,
    ws: null
};

// DOM Elements
const elements = {
    // Top HUD
    pptStatusBadge: document.getElementById('ppt-status-badge'),
    hudDeckName: document.getElementById('hud-deck-name'),
    hudSlideNum: document.getElementById('hud-slide-num'),
    hudSlideTotal: document.getElementById('hud-slide-total'),
    ttsToggleBtn: document.getElementById('tts-toggle-btn'),
    ttsStatusText: document.getElementById('tts-status-text'),

    // Sidebar & Folders
    folderPathInput: document.getElementById('folder-path-input'),
    setFolderBtn: document.getElementById('set-folder-btn'),
    refreshDecksBtn: document.getElementById('refresh-decks-btn'),
    deckListContainer: document.getElementById('deck-list-container'),

    // Remote Buttons
    btnFirst: document.getElementById('btn-first'),
    btnPrev: document.getElementById('btn-prev'),
    btnNext: document.getElementById('btn-next'),
    btnLast: document.getElementById('btn-last'),
    btnFullscreenStage: document.getElementById('btn-fullscreen-stage'),
    btnToggleFullscreen: document.getElementById('btn-toggle-fullscreen'),
    btnStopShow: document.getElementById('btn-stop-show'),
    btnBlank: document.getElementById('btn-blank'),

    // Stage Elements
    stagePanel: document.getElementById('stage-panel'),
    stageDeckTitle: document.getElementById('stage-deck-title'),
    slideCanvas: document.getElementById('slide-canvas'),
    slideContentLayout: document.getElementById('slide-content-layout'),
    slideMainTitle: document.getElementById('slide-main-title'),
    slideSubTitle: document.getElementById('slide-sub-title'),
    slideBulletList: document.getElementById('slide-bullet-list'),
    slidePagination: document.getElementById('slide-pagination'),
    slideNumberTag: document.getElementById('slide-number-tag'),
    blackoutScreen: document.getElementById('blackout-screen'),

    // Voice Hub
    micVisualizer: document.getElementById('mic-visualizer'),
    micToggleBtn: document.getElementById('mic-toggle-btn'),
    micStateTitle: document.getElementById('mic-state-title'),
    micStateSub: document.getElementById('mic-state-sub'),
    continuousMicToggle: document.getElementById('continuous-mic-toggle'),
    interimBar: document.getElementById('interim-transcript-bar'),
    interimText: document.getElementById('interim-transcript-text'),

    // Chat
    chatThread: document.getElementById('chat-thread'),
    chatForm: document.getElementById('chat-form'),
    chatInput: document.getElementById('chat-input'),
    sendBtn: document.getElementById('send-btn')
};

// Initialize on page load
document.addEventListener('DOMContentLoaded', () => {
    initSpeechRecognition();
    initWebSocket();
    fetchPresentations();
    fetchDeckSlides();
    setupEventListeners();
});

/* ==========================================================================
   Event Listeners
   ========================================================================== */
function setupEventListeners() {
    // Remote Action Buttons
    elements.btnFirst.addEventListener('click', () => sendAction('first'));
    elements.btnPrev.addEventListener('click', () => sendAction('prev'));
    elements.btnNext.addEventListener('click', () => sendAction('next'));
    elements.btnLast.addEventListener('click', () => sendAction('last'));
    elements.btnFullscreenStage.addEventListener('click', toggleFullscreenStage);
    elements.btnToggleFullscreen.addEventListener('click', toggleFullscreenStage);
    elements.btnStopShow.addEventListener('click', () => sendAction('stop_show'));
    elements.btnBlank.addEventListener('click', toggleBlackout);
    elements.blackoutScreen.addEventListener('click', toggleBlackout);

    // Folder Management
    elements.setFolderBtn.addEventListener('click', updateFolder);
    elements.refreshDecksBtn.addEventListener('click', () => fetchPresentations());

    // Mic Toggle
    elements.micToggleBtn.addEventListener('click', toggleSpeechRecognition);
    elements.continuousMicToggle.addEventListener('change', (e) => {
        state.continuous = e.target.checked;
        if (state.listening) {
            stopSpeechRecognition();
            startSpeechRecognition();
        }
    });

    // TTS Toggle
    elements.ttsToggleBtn.addEventListener('click', () => {
        state.ttsEnabled = !state.ttsEnabled;
        elements.ttsToggleBtn.classList.toggle('active', state.ttsEnabled);
        elements.ttsStatusText.textContent = state.ttsEnabled ? 'Voice Reply: ON' : 'Voice Reply: OFF';
    });

    // Keyboard Shortcuts
    window.addEventListener('keydown', (e) => {
        if (['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) return;

        if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
            e.preventDefault();
            sendAction('next');
        } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
            e.preventDefault();
            sendAction('prev');
        } else if (e.key === 'Home') {
            e.preventDefault();
            sendAction('first');
        } else if (e.key === 'End') {
            e.preventDefault();
            sendAction('last');
        } else if (e.key.toLowerCase() === 'f' || e.key === 'F5') {
            e.preventDefault();
            toggleFullscreenStage();
        } else if (e.key.toLowerCase() === 'b') {
            e.preventDefault();
            toggleBlackout();
        } else if (e.key === 'Escape') {
            if (elements.stagePanel.classList.contains('fullscreen')) {
                toggleFullscreenStage();
            }
            sendAction('stop_show');
        }
    });
}

/* ==========================================================================
   Live Presentation Stage Renderer
   ========================================================================== */
async function fetchDeckSlides(path = null) {
    try {
        const url = path ? `/api/deck/slides?path=${encodeURIComponent(path)}` : '/api/deck/slides';
        const res = await fetch(url);
        const data = await res.json();
        
        if (data.slides && data.slides.length > 0) {
            state.slides = data.slides;
            state.totalSlides = data.slides.length;
            if (data.presentation) {
                state.activeDeck = data.presentation;
                elements.hudDeckName.textContent = data.presentation;
                elements.stageDeckTitle.textContent = data.presentation;
            }
            renderSlide(state.currentSlide);
            renderPaginationDots();
        }
    } catch (e) {
        console.warn("Could not fetch deck slides:", e);
    }
}

function renderSlide(index) {
    if (!state.slides || state.slides.length === 0) return;

    const safeIndex = Math.max(1, Math.min(index, state.slides.length));
    state.currentSlide = safeIndex;
    const slide = state.slides[safeIndex - 1];

    // Update Titles
    elements.slideMainTitle.textContent = slide.title || `Slide ${safeIndex}`;
    
    if (slide.subtitle) {
        elements.slideSubTitle.textContent = slide.subtitle;
        elements.slideSubTitle.style.display = 'block';
    } else {
        elements.slideSubTitle.style.display = 'none';
    }

    // Update Bullets Content
    elements.slideBulletList.innerHTML = '';
    if (slide.content && slide.content.length > 0) {
        elements.slideBulletList.style.display = 'flex';
        slide.content.forEach(text => {
            const li = document.createElement('li');
            li.textContent = text;
            elements.slideBulletList.appendChild(li);
        });
    } else {
        elements.slideBulletList.style.display = 'none';
    }

    // Update HUD & Counters
    elements.hudSlideNum.textContent = safeIndex;
    elements.hudSlideTotal.textContent = `/ ${state.totalSlides}`;
    elements.slideNumberTag.textContent = `Slide ${safeIndex} of ${state.totalSlides}`;

    // Highlight pagination dot
    updatePaginationActiveDot(safeIndex);

    // Re-trigger CSS animation
    elements.slideContentLayout.style.animation = 'none';
    elements.slideContentLayout.offsetHeight; // trigger reflow
    elements.slideContentLayout.style.animation = null;
}

function renderPaginationDots() {
    elements.slidePagination.innerHTML = '';
    for (let i = 1; i <= state.totalSlides; i++) {
        const dot = document.createElement('div');
        dot.className = `slide-dot ${i === state.currentSlide ? 'active' : ''}`;
        dot.title = `Jump to slide ${i}`;
        dot.addEventListener('click', () => sendAction('goto', { slide: i }));
        elements.slidePagination.appendChild(dot);
    }
}

function updatePaginationActiveDot(index) {
    const dots = elements.slidePagination.querySelectorAll('.slide-dot');
    dots.forEach((dot, idx) => {
        dot.classList.toggle('active', (idx + 1) === index);
    });
}

function toggleFullscreenStage() {
    elements.stagePanel.classList.toggle('fullscreen');
    const isFull = elements.stagePanel.classList.contains('fullscreen');
    if (isFull) {
        if (elements.stagePanel.requestFullscreen) {
            elements.stagePanel.requestFullscreen().catch(() => {});
        }
    } else {
        if (document.fullscreenElement) {
            document.exitFullscreen().catch(() => {});
        }
    }
}

function toggleBlackout() {
    state.blackout = !state.blackout;
    elements.blackoutScreen.style.display = state.blackout ? 'flex' : 'none';
    sendAction('blank', { color: state.blackout ? 'black' : 'unblank' });
}

/* ==========================================================================
   Web Speech Recognition API
   ========================================================================== */
function initSpeechRecognition() {
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
        elements.micStateTitle.textContent = "Speech API Not Supported";
        elements.micStateSub.textContent = "Use Google Chrome / Edge or type in chat.";
        elements.micToggleBtn.disabled = true;
        return;
    }

    const recognition = new SpeechRecognition();
    recognition.continuous = true;
    recognition.interimResults = true;
    recognition.lang = 'en-US';

    recognition.onstart = () => {
        state.listening = true;
        elements.micVisualizer.classList.add('listening');
        elements.micStateTitle.textContent = "Listening for commands...";
        elements.micStateSub.textContent = "Say 'Next slide', 'Previous', 'Open PPT'...";
    };

    recognition.onresult = (event) => {
        let interimTranscript = '';
        let finalTranscript = '';

        for (let i = event.resultIndex; i < event.results.length; ++i) {
            if (event.results[i].isFinal) {
                finalTranscript += event.results[i][0].transcript;
            } else {
                interimTranscript += event.results[i][0].transcript;
            }
        }

        if (interimTranscript) {
            elements.interimBar.style.display = 'flex';
            elements.interimText.textContent = `Hearing: "${interimTranscript}"`;
        }

        if (finalTranscript) {
            elements.interimBar.style.display = 'none';
            const cleanText = finalTranscript.trim();
            if (cleanText) {
                handleVoiceCommand(cleanText);
            }
        }
    };

    recognition.onerror = (event) => {
        console.warn("Speech error:", event.error);
        if (event.error === 'not-allowed') {
            stopSpeechRecognition();
            appendMessage('assistant', "Please allow microphone access in your browser to use voice control.");
        }
    };

    recognition.onend = () => {
        elements.interimBar.style.display = 'none';
        if (state.listening && state.continuous) {
            try { recognition.start(); } catch (err) {}
        } else {
            state.listening = false;
            elements.micVisualizer.classList.remove('listening');
            elements.micStateTitle.textContent = "Voice Control: Standby";
            elements.micStateSub.textContent = "Click mic to start hands-free voice control";
        }
    };

    state.recognition = recognition;
}

function startSpeechRecognition() {
    if (!state.recognition) return;
    try {
        state.listening = true;
        state.recognition.start();
    } catch (e) {}
}

function stopSpeechRecognition() {
    if (!state.recognition) return;
    state.listening = false;
    try { state.recognition.stop(); } catch (e) {}
    elements.micVisualizer.classList.remove('listening');
    elements.micStateTitle.textContent = "Voice Control: Standby";
    elements.micStateSub.textContent = "Click mic to start hands-free voice control";
}

function toggleSpeechRecognition() {
    if (state.listening) {
        stopSpeechRecognition();
    } else {
        startSpeechRecognition();
    }
}

function handleVoiceCommand(transcript) {
    appendMessage('user', transcript, 'Voice');
    sendCommand(transcript, 'voice');
}

/* ==========================================================================
   Chat Submission & Actions
   ========================================================================== */
function handleChatSubmit(e) {
    e.preventDefault();
    const text = elements.chatInput.value.trim();
    if (!text) return;

    appendMessage('user', text, 'Chat');
    elements.chatInput.value = '';
    sendCommand(text, 'chat');
}

function sendQuickCommand(text) {
    appendMessage('user', text, 'Remote');
    sendCommand(text, 'chat');
}

async function sendCommand(text, source = 'chat') {
    try {
        const response = await fetch('/api/command', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ text, source })
        });
        const data = await response.json();

        if (data.feedback) {
            appendMessage('assistant', data.feedback);
            speakFeedback(data.feedback);
        }

        if (data.status) {
            updateUIStatus(data.status);
        }

        // If open command was executed, reload slide stage
        if (data.parsed && data.parsed.action === 'open') {
            fetchDeckSlides();
        }
    } catch (error) {
        appendMessage('assistant', `Error: ${error.message}`);
    }
}

async function sendAction(action, params = {}) {
    try {
        const response = await fetch('/api/action', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action, params })
        });
        const data = await response.json();
        if (data.status) {
            updateUIStatus(data.status);
        }
        if (action === 'open') {
            fetchDeckSlides();
        }
    } catch (error) {
        console.error("Action error:", error);
    }
}

async function fetchPresentations(folder = null) {
    elements.deckListContainer.innerHTML = `
        <div class="empty-state">
            <div class="spinner"></div>
            <p>Scanning presentations...</p>
        </div>
    `;

    try {
        const url = folder ? `/api/presentations?folder=${encodeURIComponent(folder)}` : '/api/presentations';
        const res = await fetch(url);
        const data = await res.json();

        state.folder = data.folder;
        elements.folderPathInput.value = data.folder;
        state.presentations = data.presentations;

        renderDeckList(data.presentations);
    } catch (error) {
        elements.deckListContainer.innerHTML = `<div class="empty-state"><p>Folder could not be loaded.</p></div>`;
    }
}

async function updateFolder() {
    const newFolder = elements.folderPathInput.value.trim();
    if (!newFolder) return;

    try {
        const res = await fetch('/api/folder', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ folder_path: newFolder })
        });
        const data = await res.json();
        if (res.ok) {
            fetchPresentations();
            fetchDeckSlides();
            appendMessage('assistant', `Loaded folder: ${data.folder}`);
        }
    } catch (err) {
        appendMessage('assistant', `Error: ${err.message}`);
    }
}

function renderDeckList(presentations) {
    if (!presentations || presentations.length === 0) {
        elements.deckListContainer.innerHTML = `
            <div class="empty-state">
                <p>No presentations in folder.</p>
                <p style="margin-top: 4px; font-size: 10px;">Add <code>.pptx</code> files to <code>${state.folder}</code></p>
            </div>
        `;
        return;
    }

    elements.deckListContainer.innerHTML = '';
    presentations.forEach((p, idx) => {
        const card = document.createElement('div');
        const num = p.index || (idx + 1);
        card.className = `deck-card ${p.is_active ? 'active' : ''}`;
        card.innerHTML = `
            <div class="deck-card-top">
                <span class="deck-number-badge">#${num}</span>
                <div class="deck-card-icon">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect>
                        <line x1="8" y1="21" x2="16" y2="21"></line>
                        <line x1="12" y1="17" x2="12" y2="21"></line>
                    </svg>
                </div>
                <div class="deck-card-meta">
                    <div class="deck-card-title" title="${p.name}">${p.name}</div>
                    <div class="deck-card-sub">${p.size_kb} KB</div>
                </div>
            </div>
            <div class="deck-card-actions">
                <button class="deck-action-btn" onclick="openDeck('${encodeURIComponent(p.path)}')">
                    ▶ Open Deck #${num}
                </button>
            </div>
        `;
        elements.deckListContainer.appendChild(card);
    });
}

function openDeck(encodedPath) {
    const fullPath = decodeURIComponent(encodedPath);
    sendAction('open', { target: fullPath, start_show: true });
    fetchDeckSlides(fullPath);
}

/* ==========================================================================
   Real-Time Status Synchronization (WebSocket)
   ========================================================================== */
function initWebSocket() {
    const protocol = window.location.protocol === 'https:' ? 'wss:' : 'ws:';
    const wsUrl = `${protocol}//${window.location.host}/ws`;

    try {
        state.ws = new WebSocket(wsUrl);
        state.ws.onmessage = (event) => {
            const msg = JSON.parse(event.data);
            if (msg.type === 'status_update') {
                updateUIStatus(msg.data);
            }
        };
        state.ws.onclose = () => setTimeout(initWebSocket, 3000);
    } catch (e) {}
}

function updateUIStatus(status) {
    if (!status) return;

    if (status.presentation_name) {
        state.activeDeck = status.presentation_name;
        elements.hudDeckName.textContent = status.presentation_name;
        elements.stageDeckTitle.textContent = status.presentation_name;
    }

    if (status.current_slide && status.current_slide !== state.currentSlide) {
        renderSlide(status.current_slide);
    }

    if (status.total_slides && status.total_slides !== state.totalSlides) {
        state.totalSlides = status.total_slides;
        renderPaginationDots();
    }
}

/* ==========================================================================
   Chat Messages & Text-to-Speech
   ========================================================================== */
function appendMessage(sender, text, sourceLabel = null) {
    const msg = document.createElement('div');
    msg.className = `message ${sender}-msg`;

    let badgeHtml = '';
    if (sourceLabel) {
        badgeHtml = `<span class="msg-badge">${sourceLabel}</span>`;
    }

    if (sender === 'user') {
        msg.innerHTML = `
            <div class="msg-content">${badgeHtml}${escapeHtml(text)}</div>
            <div class="msg-avatar">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
                    <circle cx="12" cy="7" r="4"></circle>
                </svg>
            </div>
        `;
    } else {
        msg.innerHTML = `
            <div class="msg-avatar">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon>
                </svg>
            </div>
            <div class="msg-content">${escapeHtml(text)}</div>
        `;
    }

    elements.chatThread.appendChild(msg);
    elements.chatThread.scrollTop = elements.chatThread.scrollHeight;
}

function speakFeedback(text) {
    if (!state.ttsEnabled || !window.speechSynthesis) return;
    window.speechSynthesis.cancel();
    const utterance = new SpeechSynthesisUtterance(text);
    utterance.rate = 1.05;
    window.speechSynthesis.speak(utterance);
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
