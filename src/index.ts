import { McpAgent } from "agents/mcp";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";

// Define our MCP agent with Gmail tools
export class GmailMCP extends McpAgent {
    server = new McpServer({
        name: "Gmail MCP",
        version: "0.1.0",
    });
    
    token: string | null = null;
    baseUrl = "https://small-mouse-2759.arnabbhowmik019.workers.dev";

    async init() {
        // Authentication tool
        this.server.tool(
            "authenticate",
            "Authenticate with Google to access Gmail, Calendar, and Classroom services",
            {
                token: z.string().optional().describe("The authentication token received from the auth process"),
            },
            async ({ token }) => {
                if (!token) {
                    const authUrl = `${this.baseUrl}/google/auth/gmail?redirect_url=${this.baseUrl}/token-helper`;
                    return {
                        content: [
                            {
                                type: "text",
                                text: `Please visit this URL to authorize the application:\n${authUrl}\n\nAfter authorization, you'll receive a token. Please provide that token to complete authentication.`,
                            },
                        ],
                    };
                }

                this.token = token;
                return {
                    content: [
                        {
                            type: "text",
                            text: "Authentication successful! You can now use Gmail tools.",
                        },
                    ],
                };
            }
        );

        // Search emails tool
        this.server.tool(
            "search_emails",
            "Search for emails in Gmail using Gmail search syntax",
            {
                query: z.string().describe("Gmail search query (e.g., \"from:example@gmail.com\", \"subject:important\", \"is:unread\")"),
            },
            async ({ query }) => {
                try {
                    const result = await this.makeRequest(`/gmail/search?q=${encodeURIComponent(query)}`);
                    
                    if (!result?.messages?.length) {
                        return {
                            content: [
                                {
                                    type: "text",
                                    text: "No messages found matching the search query.",
                                },
                            ],
                        };
                    }

                    const messageList = result.messages.map((msg: any) => 
                        `ID: ${msg.id}\n` +
                        `Thread ID: ${msg.threadId}\n` +
                        (msg.snippet ? `Snippet: ${msg.snippet}\n` : '') +
                        '---'
                    ).join('\n');

                    return {
                        content: [
                            {
                                type: "text",
                                text: `Found ${result.messages.length} messages:\n\n${messageList}`,
                            },
                        ],
                    };
                } catch (error) {
                    return {
                        content: [
                            {
                                type: "text",
                                text: `Error: ${(error as Error).message}`,
                            },
                        ],
                    };
                }
            }
        );

        // Get email tool
        this.server.tool(
            "get_email",
            "Retrieve a specific email by its message ID",
            {
                messageId: z.string().describe("The Gmail message ID"),
            },
            async ({ messageId }) => {
                try {
                    const message = await this.makeRequest(`/gmail/message?messageId=${messageId}&decode=true`);
                    
                    let content = `Message ID: ${message.id}\n`;
                    if (message.threadId) content += `Thread ID: ${message.threadId}\n`;
                    if (message.snippet) content += `Snippet: ${message.snippet}\n\n`;

                    if (message.payload?.headers) {
                        content += 'Headers:\n';
                        message.payload.headers.forEach((header: any) => {
                            content += `${header.name}: ${header.value}\n`;
                        });
                    }

                    return {
                        content: [
                            {
                                type: "text",
                                text: content,
                            },
                        ],
                    };
                } catch (error) {
                    return {
                        content: [
                            {
                                type: "text",
                                text: `Error: ${(error as Error).message}`,
                            },
                        ],
                    };
                }
            }
        );

        // List labels tool
        this.server.tool(
            "list_labels",
            "List all Gmail labels in the user's account",
            {},
            async () => {
                try {
                    const result = await this.makeRequest('/gmail/labels');
                    
                    if (!result?.labels?.length) {
                        return {
                            content: [
                                {
                                    type: "text",
                                    text: "No labels found.",
                                },
                            ],
                        };
                    }

                    const labelText = result.labels
                        .map((label: any) => `${label.name} (${label.id}) - ${label.type}`)
                        .join('\n');

                    return {
                        content: [
                            {
                                type: "text",
                                text: `Gmail Labels:\n\n${labelText}`,
                            },
                        ],
                    };
                } catch (error) {
                    return {
                        content: [
                            {
                                type: "text",
                                text: `Error: ${(error as Error).message}`,
                            },
                        ],
                    };
                }
            }
        );

        // Get profile tool
        this.server.tool(
            "get_profile",
            "Get information about the user's Gmail profile",
            {},
            async () => {
                try {
                    const profile = await this.makeRequest('/gmail/list');
                    
                    return {
                        content: [
                            {
                                type: "text",
                                text: `Gmail Profile:\n${JSON.stringify(profile, null, 2)}`,
                            },
                        ],
                    };
                } catch (error) {
                    return {
                        content: [
                            {
                                type: "text",
                                text: `Error: ${(error as Error).message}`,
                            },
                        ],
                    };
                }
            }
        );

        // Calendar events tools
        this.server.tool(
            "list_events",
            "List calendar events within a specified time range",
            {
                timeMin: z.string().optional().describe("Start time for listing events (RFC3339 timestamp)"),
                timeMax: z.string().optional().describe("End time for listing events (RFC3339 timestamp)"),
                maxResults: z.number().optional().describe("Maximum number of events to return"),
            },
            async ({ timeMin, timeMax, maxResults }) => {
                try {
                    const queryParams = new URLSearchParams();
                    if (timeMin) queryParams.append('timeMin', timeMin);
                    if (timeMax) queryParams.append('timeMax', timeMax);
                    if (maxResults) queryParams.append('maxResults', maxResults.toString());
                    
                    const events = await this.makeRequest(`/calendar/events?${queryParams.toString()}`);
                    
                    if (!events?.items?.length) {
                        return {
                            content: [
                                {
                                    type: "text",
                                    text: "No events found.",
                                },
                            ],
                        };
                    }

                    const eventList = events.items.map((event: any) => {
                        let details = `Title: ${event.summary}\n`;
                        if (event.description) details += `Description: ${event.description}\n`;
                        details += `Start: ${event.start.dateTime || event.start.date}\n`;
                        details += `End: ${event.end.dateTime || event.end.date}\n`;
                        if (event.attendees?.length) {
                            details += 'Attendees:\n';
                            event.attendees.forEach((attendee: any) => {
                                details += `  - ${attendee.email}\n`;
                            });
                        }
                        details += '---\n';
                        return details;
                    }).join('\n');

                    return {
                        content: [
                            {
                                type: "text",
                                text: `Found ${events.items.length} events:\n\n${eventList}`,
                            },
                        ],
                    };
                } catch (error) {
                    return {
                        content: [
                            {
                                type: "text",
                                text: `Error: ${(error as Error).message}`,
                            },
                        ],
                    };
                }
            }
        );

        // Create calendar event
        this.server.tool(
            "create_event",
            "Create a new event in Google Calendar",
            {
                summary: z.string().describe("Title of the event"),
                description: z.string().optional().describe("Description of the event"),
                start: z.object({
                    dateTime: z.string().describe("Start time (RFC3339 timestamp)"),
                    timeZone: z.string().optional().describe("Timezone for the start time"),
                }),
                end: z.object({
                    dateTime: z.string().describe("End time (RFC3339 timestamp)"),
                    timeZone: z.string().optional().describe("Timezone for the end time"),
                }),
                attendees: z
                    .array(
                        z.object({
                            email: z.string().describe("Email address of the attendee"),
                        })
                    )
                    .optional()
                    .describe("List of attendees"),
            },
            async (args) => {
                try {
                    const response = await this.makeRequest('/calendar/events', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                        },
                        body: JSON.stringify(args),
                    });

                    return {
                        content: [
                            {
                                type: "text",
                                text: `Event created successfully!\nID: ${response.id}\nTitle: ${response.summary}\nStart: ${response.start.dateTime}\nEnd: ${response.end.dateTime}`,
                            },
                        ],
                    };
                } catch (error) {
                    return {
                        content: [
                            {
                                type: "text",
                                text: `Error: ${(error as Error).message}`,
                            },
                        ],
                    };
                }
            }
        );

        // Classroom tools
        this.server.tool(
            "list_courses",
            "List all Google Classroom courses available to the user",
            {},
            async () => {
                try {
                    const courses = await this.makeRequest('/classroom/courses');
                    
                    if (!courses?.courses?.length) {
                        return {
                            content: [
                                {
                                    type: "text",
                                    text: "No courses found.",
                                },
                            ],
                        };
                    }

                    const courseList = courses.courses.map((course: any) => {
                        let details = `Name: ${course.name}\n`;
                        details += `ID: ${course.id}\n`;
                        if (course.section) details += `Section: ${course.section}\n`;
                        if (course.description) details += `Description: ${course.description}\n`;
                        details += `---\n`;
                        return details;
                    }).join('\n');

                    return {
                        content: [
                            {
                                type: "text",
                                text: `Found ${courses.courses.length} courses:\n\n${courseList}`,
                            },
                        ],
                    };
                } catch (error) {
                    return {
                        content: [
                            {
                                type: "text",
                                text: `Error: ${(error as Error).message}`,
                            },
                        ],
                    };
                }
            }
        );

        // List coursework tool
        this.server.tool(
            "list_coursework",
            "List coursework for a specific Google Classroom course",
            {
                courseId: z.string().describe("The ID of the course to retrieve coursework from"),
            },
            async ({ courseId }) => {
                try {
                    const coursework = await this.makeRequest(`/classroom/coursework?courseId=${courseId}`);
                    
                    if (!coursework?.courseWork?.length) {
                        return {
                            content: [
                                {
                                    type: "text",
                                    text: "No coursework found for this course."
                                }
                            ]
                        };
                }

                const courseworkList = coursework.courseWork.map((work: any) => {
                    let details = `Title: ${work.title}\n`;
                    details += `ID: ${work.id}\n`;
                    if (work.description) details += `Description: ${work.description}\n`;
                    if (work.dueDate) {
                        const dueDate = work.dueDate;
                        const dueTime = work.dueTime || { hours: 0, minutes: 0 };
                        details += `Due: ${dueDate.year}-${dueDate.month}-${dueDate.day} ${dueTime.hours}:${dueTime.minutes}\n`;
                    }
                    if (work.maxPoints) details += `Max Points: ${work.maxPoints}\n`;
                    details += `---\n`;
                    return details;
                }).join('\n');

                return {
                    content: [
                        {
                            type: "text",
                            text: `Found ${coursework.courseWork.length} coursework items:\n\n${courseworkList}`,
                        },
                    ],
                };
            } catch (error) {
                return {
                    content: [
                        {
                            type: "text",
                            text: `Error: ${(error as Error).message}`,
                        },
                    ],
                };
            }
        );

        // List announcements tool
        this.server.tool(
            "list_announcements",
            "List announcements for a specific Google Classroom course",
            {
                courseId: z.string().describe("The ID of the course to retrieve announcements from"),
            },
            async ({ courseId }) => {
                try {
                    const announcements = await this.makeRequest(`/classroom/announcements?courseId=${courseId}`);
                    
                    if (!announcements?.announcements?.length) {
                        return {
                            content: [
                                {
                                    type: "text",
                                    text: "No announcements found for this course."
                                }
                            ]
                        };
                }

                const announcementList = announcements.announcements.map((announcement: any) => {
                    let details = `ID: ${announcement.id}\n`;
                    details += `Posted: ${new Date(announcement.creationTime).toLocaleString()}\n`;
                    details += `Text: ${announcement.text}\n`;
                    if (announcement.materials?.length) {
                        details += 'Materials:\n';
                        announcement.materials.forEach((material: any) => {
                            if (material.driveFile) details += `  - Drive File: ${material.driveFile.title}\n`;
                            if (material.link) details += `  - Link: ${material.link.url} (${material.link.title})\n`;
                            if (material.youtubeVideo) details += `  - YouTube: ${material.youtubeVideo.title}\n`;
                        });
                    }
                    details += `---\n`;
                    return details;
                }).join('\n');

                return {
                    content: [
                        {
                            type: "text",
                            text: `Found ${announcements.announcements.length} announcements:\n\n${announcementList}`,
                        },
                    ],
                };
            } catch (error) {
                return {
                    content: [
                        {
                            type: "text",
                            text: `Error: ${(error as Error).message}`,
                        },
                    ],
                };
            }
        );

        // Get coursework details tool
        this.server.tool(
            "get_coursework",
            "Get detailed content of a specific coursework/assignment",
            {
                courseId: z.string().describe("The ID of the course containing the assignment"),
                courseworkId: z.string().describe("The ID of the specific coursework/assignment"),
            },
            async ({ courseId, courseworkId }) => {
                try {
                    const coursework = await this.makeRequest(`/classroom/coursework/details?courseId=${courseId}&courseworkId=${courseworkId}`);
                    
                    if (!coursework) {
                        return {
                            content: [
                                {
                                    type: "text",
                                    text: "Assignment not found or could not be accessed."
                                }
                            ]
                        };
                }

                let details = `Title: ${coursework.title}\n`;
                details += `ID: ${coursework.id}\n`;
                details += `Course ID: ${coursework.courseId}\n`;
                
                if (coursework.description) details += `Description: ${coursework.description}\n`;
                if (coursework.state) details += `State: ${coursework.state}\n`;
                if (coursework.creationTime) details += `Created: ${new Date(coursework.creationTime).toLocaleString()}\n`;
                if (coursework.updateTime) details += `Updated: ${new Date(coursework.updateTime).toLocaleString()}\n`;
                
                if (coursework.dueDate) {
                    const dueDate = coursework.dueDate;
                    const dueTime = coursework.dueTime || { hours: 0, minutes: 0 };
                    details += `Due: ${dueDate.year}-${dueDate.month}-${dueDate.day} ${dueTime.hours}:${dueTime.minutes}\n`;
                }
                
                if (coursework.maxPoints) details += `Max Points: ${coursework.maxPoints}\n`;
                if (coursework.workType) details += `Work Type: ${coursework.workType}\n`;
                
                if (coursework.materials?.length) {
                    details += 'Materials:\n';
                    coursework.materials.forEach((material: any) => {
                        if (material.driveFile) {
                            // Handle different possible structures
                            const fileData = material.driveFile.driveFile || material.driveFile;
                            details += `  - Drive File: ${fileData.title || fileData.name}\n`;
                            details += `    ID: ${fileData.id}\n`;
                            details += `    Link: ${fileData.alternateLink || fileData.webViewLink}\n`;
                        }
                        if (material.link) {
                            details += `  - Link: ${material.link.url}\n`;
                            if (material.link.title) details += `    Title: ${material.link.title}\n`;
                        }
                        if (material.youtubeVideo) {
                            details += `  - YouTube: ${material.youtubeVideo.title}\n`;
                            details += `    Link: ${material.youtubeVideo.alternateLink}\n`;
                        }
                        if (material.form) {
                            details += `  - Form: ${material.form.title}\n`;
                            details += `    Link: ${material.form.formUrl}\n`;
                        }
                    });
                }

                return {
                    content: [
                        {
                            type: "text",
                            text: details,
                        },
                    ],
                };
            } catch (error) {
                return {
                    content: [
                        {
                            type: "text",
                            text: `Error: ${(error as Error).message}`,
                        },
                    ],
                };
            }
        );

        // Get assignment materials tool
        this.server.tool(
            "get_assignment_materials",
            "Get direct access to files attached to a classroom assignment",
            {
                courseId: z.string().describe("The ID of the course containing the assignment"),
                courseworkId: z.string().describe("The ID of the specific coursework/assignment"),
            },
            async ({ courseId, courseworkId }) => {
                try {
                    const coursework = await this.makeRequest(`/classroom/coursework/details?courseId=${courseId}&courseworkId=${courseworkId}`);
                    
                    if (!coursework || !coursework.materials?.length) {
                        return {
                            content: [
                                {
                                    type: "text",
                                    text: "No materials found for this assignment.",
                            },
                        ],
                    };
                }

                // Transform materials into a more accessible format
                const materials: any[] = [];
                
                coursework.materials.forEach((material: any) => {
                    if (material.driveFile) {
                        // Handle different possible structures
                        const fileData = material.driveFile.driveFile || material.driveFile;
                        materials.push({
                            type: 'drive_file',
                            id: fileData.id,
                            title: fileData.title || fileData.name,
                            url: fileData.alternateLink || fileData.webViewLink,
                            mimeType: fileData.mimeType
                        });
                    }
                    if (material.link) {
                        materials.push({
                            type: 'link',
                            url: material.link.url,
                            title: material.link.title || material.link.url
                        });
                    }
                    if (material.youtubeVideo) {
                        materials.push({
                            type: 'youtube',
                            id: material.youtubeVideo.id,
                            title: material.youtubeVideo.title,
                            url: material.youtubeVideo.alternateLink
                        });
                    }
                    if (material.form) {
                        materials.push({
                            type: 'form',
                            id: material.form.id,
                            title: material.form.title,
                            url: material.form.formUrl
                        });
                    }
                });

                return {
                    content: [
                        {
                            type: "text",
                            text: `Assignment Materials for "${coursework.title}":\n\n${JSON.stringify(materials, null, 2)}`,
                        },
                    ],
                };
            } catch (error) {
                return {
                    content: [
                        {
                            type: "text",
                            text: `Error: ${(error as Error).message}`,
                        },
                    ],
                };
            }
        );

        // Download file tool
        this.server.tool(
            "download_file",
            "Download a document, presentation, or other file",
            {
                fileId: z.string().describe("The Drive file ID to download"),
                format: z.enum(['text', 'html', 'raw']).optional().describe("Optional format for conversion (e.g., \"text\", \"html\")"),
            },
            async ({ fileId, format = 'text' }) => {
                try {
                    const fileContent = await this.makeRequest(`/drive/files/download?fileId=${fileId}&format=${format}`);
                    
                    if (!fileContent || !fileContent.content) {
                        // If content is missing, try to get at least file metadata
                        const fileMetadata = await this.makeRequest(`/drive/files?fileId=${fileId}`);
                        return {
                            content: [
                                {
                                    type: "text",
                                    text: `File metadata retrieved:\n${JSON.stringify(fileMetadata, null, 2)}\n\nTo view this file, use the web link above or try the read_document tool with different fileType options.`,
                            },
                        ],
                    };
                }
                
                return {
                    content: [
                        {
                            type: "text",
                            text: fileContent.content || 'File downloaded successfully, but content is empty or binary.',
                        },
                    ],
                };
            } catch (error) {
                return {
                    content: [
                        {
                            type: "text",
                            text: `Failed to download file: ${(error as Error).message}\n\nFor PowerPoint files, try using the read_document tool with fileType='slides' or visit the file's web link directly.`,
                        },
                    ],
                };
            }
        );

        // Read document tool
        this.server.tool(
            "read_document",
            "Process and read content from a document file",
            {
                fileId: z.string().describe("The Drive file ID to read"),
                fileType: z.enum(['doc', 'pdf', 'slides', 'sheet', 'auto']).optional().describe("The type of file (e.g., \"doc\", \"pdf\", \"slides\")"),
            },
            async ({ fileId, fileType = 'auto' }) => {
                try {
                    const documentContent = await this.makeRequest(`/drive/files/read?fileId=${fileId}&fileType=${fileType}`);
                    
                    if (!documentContent || !documentContent.content) {
                        return {
                            content: [
                                {
                                    type: "text",
                                    text: 'Could not extract readable content from this document.',
                            },
                        ],
                    };
                }
                
                return {
                    content: [
                        {
                            type: "text",
                            text: `Document Title: ${documentContent.title || 'Unknown'}\n\n${documentContent.content}`,
                        },
                    ],
                };
            } catch (error) {
                return {
                    content: [
                        {
                            type: "text",
                            text: `Failed to read document: ${(error as Error).message}`,
                        },
                    ],
                };
            }
        );

        // Add more tools based on the provided Node.js implementation
    }

    // Helper method to make authenticated requests
    async makeRequest(endpoint: string, options: RequestInit = {}) {
        if (!this.token) {
            throw new Error('Authentication required. Please use the authenticate tool first with no parameters to get the authorization URL. After visiting that URL and completing authentication, call authenticate again with the token you receive.');
        }

        const response = await fetch(`${this.baseUrl}${endpoint}`, {
            ...options,
            headers: {
                ...options.headers,
                'Authorization': `Bearer ${this.token}`
            }
        });

        if (!response.ok) {
            const errorData = await response.text();
            throw new Error(`API request failed (${response.status}): ${response.statusText}. ${errorData}`);
        }

        return response.json();
    }
}

// Export GmailMCP as a Durable Object
export { GmailMCP as MyMCP };

// Interface for Cloudflare Worker environment
interface Env {
    // Add any Worker environment bindings here
    MyMCP: DurableObjectNamespace;
}

export default {
    fetch(request: Request, env: Env, ctx: ExecutionContext) {
        const url = new URL(request.url);

        if (url.pathname === "/sse" || url.pathname === "/sse/message") {
            return GmailMCP.serveSSE("/sse").fetch(request, env, ctx);
        }

        if (url.pathname === "/mcp" || url.pathname === "/") {
            return GmailMCP.serve("/mcp").fetch(request, env, ctx);
        }

        return new Response("Not found", { status: 404 });
    },
};
