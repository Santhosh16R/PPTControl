/**
 * OpenAI Voice / Realtime Agent Integration for Remote PowerPoint Controller
 * 
 * Instructions:
 * 1. Place this file or copy the tools/handlers into your OpenAI Voice Agent project.
 * 2. Set PRESENTATION_PC_URL to the IP address of the PC running the PowerPoint server (e.g. 'http://192.168.1.15:3000').
 */

// Replace with your Presentation PC's local IP address
const PRESENTATION_PC_URL = process.env.PPT_SERVER_URL || 'http://192.168.1.15:3000';

/**
 * 1. OpenAI Function Tool Definitions
 * Pass this 'tools' array in your OpenAI session / model config
 */
const pptTools = [
    {
        type: "function",
        name: "open_presentation",
        description: "Open a PowerPoint presentation on the screen. Can specify presentation name, index number (1, 2, 3), or ordinals ('first', 'second', '3rd', 'last').",
        parameters: {
            type: "object",
            properties: {
                target: {
                    type: "string",
                    description: "The name, number, or ordinal of the presentation to open (e.g., '1', '2', 'first', 'second', 'sample', 'last')."
                }
            }
        }
    },
    {
        type: "function",
        name: "next_slide",
        description: "Advance to the next slide in the active PowerPoint presentation.",
        parameters: {
            type: "object",
            properties: {}
        }
    },
    {
        type: "function",
        name: "previous_slide",
        description: "Go back to the previous slide in the active PowerPoint presentation.",
        parameters: {
            type: "object",
            properties: {}
        }
    },
    {
        type: "function",
        name: "goto_slide",
        description: "Jump directly to a specific slide number in the active presentation.",
        parameters: {
            type: "object",
            properties: {
                slide_number: {
                    type: "integer",
                    description: "The slide number to jump to (1-based index)."
                }
            },
            required: ["slide_number"]
        }
    },
    {
        type: "function",
        name: "list_presentations",
        description: "List all available PowerPoint presentation files (.pptx / .ppt) found on the presentation PC.",
        parameters: {
            type: "object",
            properties: {}
        }
    },
    {
        type: "function",
        name: "start_slideshow",
        description: "Start or resume fullscreen presentation mode.",
        parameters: {
            type: "object",
            properties: {}
        }
    },
    {
        type: "function",
        name: "stop_slideshow",
        description: "Exit or stop the fullscreen presentation mode.",
        parameters: {
            type: "object",
            properties: {}
        }
    },
    {
        type: "function",
        name: "toggle_blackout",
        description: "Toggle screen blackout on/off during a presentation.",
        parameters: {
            type: "object",
            properties: {}
        }
    }
];

/**
 * 2. Remote API Dispatcher
 * Calls the Presentation PC HTTP API across the local network
 */
async function callPresentationServer(action, params = {}) {
    try {
        const response = await fetch(`${PRESENTATION_PC_URL}/api/action`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action, params })
        });
        const data = await response.json();
        return data;
    } catch (error) {
        console.error(`[PPT Remote Error] Failed to execute '${action}':`, error.message);
        return { success: false, error: `Could not connect to Presentation PC at ${PRESENTATION_PC_URL}. Verify it is running.` };
    }
}

/**
 * 3. OpenAI Tool Call Handler
 * Call this function whenever OpenAI returns a function call
 */
async function handleOpenAIToolCall(toolName, args = {}) {
    console.log(`[OpenAI Voice Agent] Executing tool: ${toolName}`, args);

    switch (toolName) {
        case "open_presentation":
            return await callPresentationServer('open', { target: args.target });

        case "next_slide":
            return await callPresentationServer('next');

        case "previous_slide":
            return await callPresentationServer('prev');

        case "goto_slide":
            return await callPresentationServer('goto', { slide: args.slide_number });

        case "list_presentations":
            try {
                const res = await fetch(`${PRESENTATION_PC_URL}/api/presentations`);
                const data = await res.json();
                return data;
            } catch (err) {
                return { success: false, error: err.message };
            }

        case "start_slideshow":
            return await callPresentationServer('start_show');

        case "stop_slideshow":
            return await callPresentationServer('stop_show');

        case "toggle_blackout":
            return await callPresentationServer('blank');

        default:
            return { error: `Unknown tool name: ${toolName}` };
    }
}

module.exports = {
    pptTools,
    handleOpenAIToolCall,
    callPresentationServer
};
