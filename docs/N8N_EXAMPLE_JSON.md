# n8n quick setup template
  Copy the below json in your clipboard. And paste to your n8n canvas.
  
{
  "nodes": [
    {
      "parameters": {
        "jsCode": "// Process HTTP MCP response\n   const response = items[0].json;\n   \n   if (response.result && response.result.folders && Array.isArray(response.result.folders)) {\n     const folders = response.result.folders;\n     return [{\n       json: {\n         success: true,\n         message: \"MCP server connection successful (HTTP)\",\n         folderCount: folders.length,\n         folders: folders.map(f => f.name || f.display_name || 'Unknown'),\n         folderDetails: folders.map(f => ({\n           name: f.name || f.display_name,\n           displayName: f.display_name,\n           fullPath: f.full_path,\n           itemCount: f.item_count,\n           unreadCount: f.unread_count,\n           folderType: f.folder_type,\n           accessible: f.accessible,\n           hasSubfolders: f.has_subfolders\n         })),\n         mode: \"http\",\n         raw_response: response\n       }\n     }];\n   } else if (response.error) {\n     return [{\n       json: {\n         success: false,\n         message: \"MCP server error\",\n         error: response.error.message,\n         error_code: response.error.code,\n         mode: \"http\"\n       }\n     }];\n   } else {\n     return [{\n       json: {\n         success: false,\n         message: \"Unexpected response format\",\n         response: response,\n         mode: \"http\"\n       }\n     }];\n   }"
      },
      "type": "n8n-nodes-base.code",
      "typeVersion": 2,
      "position": [
        656,
        -80
      ],
      "id": "7a7035b7-cdc6-4880-b4fe-c3003917f959",
      "name": "Code in JavaScript"
    },
    {
      "parameters": {
        "url": "http://192.168.1.164:8080/health",
        "options": {}
      },
      "type": "n8n-nodes-base.httpRequest",
      "typeVersion": 4.2,
      "position": [
        16,
        0
      ],
      "id": "37f1384d-4212-42a5-9cf8-7d54a69d120d",
      "name": "MCP Server health check"
    },
    {
      "parameters": {
        "errorMessage": "MCP server is not in healthy status. Workflow terminated."
      },
      "type": "n8n-nodes-base.stopAndError",
      "typeVersion": 1,
      "position": [
        448,
        112
      ],
      "id": "2c845bc4-c942-4cf3-b7d1-8c61c7fdb9d1",
      "name": "Stop and Error"
    },
    {
      "parameters": {
        "conditions": {
          "options": {
            "caseSensitive": false,
            "leftValue": "",
            "typeValidation": "strict",
            "version": 2
          },
          "conditions": [
            {
              "id": "7a2b89b6-7fa5-4fec-a5f5-7cc05665a58a",
              "leftValue": "={{ $json.status }}",
              "rightValue": "healthy",
              "operator": {
                "type": "string",
                "operation": "equals",
                "name": "filter.operator.equals"
              }
            }
          ],
          "combinator": "and"
        },
        "options": {
          "ignoreCase": true
        }
      },
      "type": "n8n-nodes-base.if",
      "typeVersion": 2.2,
      "position": [
        224,
        0
      ],
      "id": "7a390c18-ef08-43f2-8685-5122aaf030ab",
      "name": "MCP Server alive"
    },
    {
      "parameters": {
        "method": "POST",
        "url": "http://192.168.1.164:8080/mcp",
        "sendHeaders": true,
        "headerParameters": {
          "parameters": [
            {
              "name": "Content-Type",
              "value": "application/json"
            }
          ]
        },
        "sendBody": true,
        "specifyBody": "json",
        "jsonBody": "{\n     \"jsonrpc\": \"2.0\",\n     \"id\": \"test-folders\",\n     \"method\": \"get_folders\",\n     \"params\": {}\n   }",
        "options": {
          "timeout": 10000
        }
      },
      "type": "n8n-nodes-base.httpRequest",
      "typeVersion": 4.2,
      "position": [
        448,
        -80
      ],
      "id": "7441c6f9-31ad-4e89-a4b0-b61b91fca5d6",
      "name": "POST /mcp /get_folders",
      "alwaysOutputData": true,
      "retryOnFail": true,
      "notes": "POST request to /mcp"
    },
    {
      "parameters": {
        "language": "python",
        "pythonCode": "# Get the folders array (simple list of folder names)\nfolders = _input.first()['json']['folders']\n\n# Check if '收件匣' is in the folders\nif '收件匣' in folders:\n    inbox_name = '收件匣'\n    return [{'json': {'inbox_name': inbox_name}}]\nelse:\n    return [{'json': {'inbox_name': None}}]"
      },
      "type": "n8n-nodes-base.code",
      "typeVersion": 2,
      "position": [
        1088,
        -160
      ],
      "id": "991663b8-b0d0-452a-aaf3-10e6fda457ee",
      "name": "Code in Python (Beta)"
    },
    {
      "parameters": {
        "conditions": {
          "options": {
            "caseSensitive": true,
            "leftValue": "",
            "typeValidation": "strict",
            "version": 2
          },
          "conditions": [
            {
              "id": "8be22664-0eb7-42b8-94e6-67b77f66545b",
              "leftValue": "={{ $json.success }}",
              "rightValue": "=",
              "operator": {
                "type": "boolean",
                "operation": "true",
                "singleValue": true
              }
            },
            {
              "id": "2b891ab0-c3b6-408a-a7e0-029b360b9f9e",
              "leftValue": "={{ $json.folders }}",
              "rightValue": "inbox",
              "operator": {
                "type": "array",
                "operation": "contains",
                "rightType": "any"
              }
            }
          ],
          "combinator": "and"
        },
        "options": {}
      },
      "type": "n8n-nodes-base.if",
      "typeVersion": 2.2,
      "position": [
        864,
        48
      ],
      "id": "a895c5f1-edd0-4c5e-ac51-0cbc30af8a22",
      "name": "if 'inbox' in folders",
      "alwaysOutputData": false
    },
    {
      "parameters": {
        "conditions": {
          "options": {
            "caseSensitive": true,
            "leftValue": "",
            "typeValidation": "strict",
            "version": 2
          },
          "conditions": [
            {
              "id": "8be22664-0eb7-42b8-94e6-67b77f66545b",
              "leftValue": "={{ $json.success }}",
              "rightValue": "=",
              "operator": {
                "type": "boolean",
                "operation": "true",
                "singleValue": true
              }
            },
            {
              "id": "2b891ab0-c3b6-408a-a7e0-029b360b9f9e",
              "leftValue": "={{ $json.folders }}",
              "rightValue": "收件匣",
              "operator": {
                "type": "array",
                "operation": "contains",
                "rightType": "any"
              }
            }
          ],
          "combinator": "and"
        },
        "options": {}
      },
      "type": "n8n-nodes-base.if",
      "typeVersion": 2.2,
      "position": [
        864,
        -144
      ],
      "id": "2d31ac98-b00a-4541-a412-bd281c3a4c40",
      "name": "if '收件匣' in folders",
      "alwaysOutputData": false
    },
    {
      "parameters": {
        "language": "python",
        "pythonCode": "# Get the folders array (simple list of folder names)\nfolders = _input.first()['json']['folders']\n\n# Check if 'inbox' is in the folders\nif 'inbox' in folders:\n    inbox_name = 'inbox'\n    return [{'json': {'inbox_name': inbox_name}}]\nelse:\n    return [{'json': {'inbox_name': None}}]"
      },
      "type": "n8n-nodes-base.code",
      "typeVersion": 2,
      "position": [
        1088,
        32
      ],
      "id": "de4ad46c-0959-46d2-b018-88baa5302f85",
      "name": "Code in Python (Beta)1"
    },
    {
      "parameters": {
        "modelId": {
          "__rl": true,
          "value": "models/gemini-2.5-flash",
          "mode": "list",
          "cachedResultName": "models/gemini-2.5-flash"
        },
        "messages": {
          "values": [
            {
              "content": "You are a very experienced software R&D in personal computer manufacturing. Your department responsibility is to analyze and debug software related issues. Such as driver, operating system, apps related. There are several departments will participant in the project, such as hardware(EE), BIOS(firmware), mechenical(chasis), and other keyparts(LCD panel, keyboard, camera,...etc). You'll need to distinquish which team is most likely to be responsible for priliminary debug.",
              "role": "model"
            },
            {
              "content": "Here's a bug report email. Look into the email and analyze the issue report to see if software team should take a look first than other teams.\n\nIssue report from email:\n"
            }
          ]
        },
        "options": {}
      },
      "type": "@n8n/n8n-nodes-langchain.googleGemini",
      "typeVersion": 1,
      "position": [
        2288,
        -720
      ],
      "id": "1766b594-fb8d-42b4-9adb-ed5b09bcaa4a",
      "name": "Message a model",
      "credentials": {
        "googlePalmApi": {
          "id": "U6Pcb79nIRa1KZD2",
          "name": "Google Gemini(PaLM) Api account"
        }
      }
    },
    {
      "parameters": {
        "toolDescription": "Get unread email from inbox. ",
        "method": "POST",
        "url": "http://192.168.1.164:8080/mcp",
        "sendHeaders": true,
        "headerParameters": {
          "parameters": [
            {
              "name": "Content-Type",
              "value": "application/json"
            }
          ]
        },
        "sendBody": true,
        "specifyBody": "json",
        "jsonBody": "{\n  \"jsonrpc\": \"2.0\",\n  \"id\": \"list-inbox-emails-request\",\n  \"method\": \"list_inbox_emails\",\n  \"params\": {\n    \"limit\": 10,\n    \"unread_only\": true\n  }\n}",
        "options": {}
      },
      "type": "n8n-nodes-base.httpRequestTool",
      "typeVersion": 4.2,
      "position": [
        2608,
        -384
      ],
      "id": "a0353f3b-0a2f-4968-ac83-9f4724059a1c",
      "name": "Get unread emails"
    },
    {
      "parameters": {
        "rule": {
          "interval": [
            {
              "triggerAtHour": 8
            }
          ]
        }
      },
      "type": "n8n-nodes-base.scheduleTrigger",
      "typeVersion": 1.2,
      "position": [
        -176,
        0
      ],
      "id": "cb62eadd-72b9-47b0-854e-924022f234ab",
      "name": "Trigger at 8am everyday"
    },
    {
      "parameters": {
        "toolDescription": "search email",
        "method": "POST",
        "url": "http://192.168.1.164:8080/mcp",
        "sendHeaders": true,
        "headerParameters": {
          "parameters": [
            {
              "name": "Content-Type",
              "value": "application/json"
            }
          ]
        },
        "sendBody": true,
        "specifyBody": "json",
        "jsonBody": "{\n  \"jsonrpc\": \"2.0\",\n  \"id\": \"search-folder\",\n  \"method\": \"search_emails\",\n  \"params\": {\n    \"query\": \"subject:Summary\",\n    \"folder_id\": \"00000000DB2820C5F3F8204492F273035529BA6801009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000\",\n    \"limit\": 10\n  }\n}",
        "options": {}
      },
      "type": "n8n-nodes-base.httpRequestTool",
      "typeVersion": 4.2,
      "position": [
        2288,
        -384
      ],
      "id": "fc163823-063d-4027-ae80-9fd3325e2aad",
      "name": "search_emails"
    },
    {
      "parameters": {
        "toolDescription": "Send mail",
        "method": "POST",
        "url": "http://192.168.1.164:8080/mcp",
        "sendHeaders": true,
        "headerParameters": {
          "parameters": [
            {
              "name": "Content-Type",
              "value": "application/json"
            }
          ]
        },
        "sendBody": true,
        "specifyBody": "json",
        "jsonBody": "{\n  \"jsonrpc\": \"2.0\",\n  \"id\": \"send-email-working\",\n  \"method\": \"send_email\",\n  \"params\": {\n    \"to_recipients\": [\"JackieCF_Lin@compal.com\"],\n    \"subject\": \"Automated Report from N8N\",\n    \"body\": \"This email was sent automatically from n8n workflow 2359.\",\n    \"body_format\": \"html\"\n  }\n}",
        "options": {}
      },
      "type": "n8n-nodes-base.httpRequestTool",
      "typeVersion": 4.2,
      "position": [
        2288,
        -560
      ],
      "id": "d4392d0f-e6ed-420c-a928-b7ac5d768928",
      "name": "send_mails"
    },
    {
      "parameters": {
        "toolDescription": "get email",
        "method": "POST",
        "url": "http://192.168.1.164:8080/mcp",
        "sendHeaders": true,
        "headerParameters": {
          "parameters": [
            {
              "name": "Content-Type",
              "value": "application/json"
            }
          ]
        },
        "sendBody": true,
        "specifyBody": "json",
        "jsonBody": "{\n  \"jsonrpc\": \"2.0\",\n  \"id\": \"get-email-request\",\n  \"method\": \"get_email\",\n  \"params\": {\n    \"email_id\": \"00000000DB2820C5F3F8204492F273035529BA6807009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000C8F3FCAAA615D740A192DA27F69F25450000365C4DE70000\"\n  }\n}\n\n",
        "options": {}
      },
      "type": "n8n-nodes-base.httpRequestTool",
      "typeVersion": 4.2,
      "position": [
        2592,
        -560
      ],
      "id": "2be63cfd-8798-47af-b9c1-a5161fde0b6a",
      "name": "get email(id)"
    },
    {
      "parameters": {
        "toolDescription": "list email",
        "method": "POST",
        "url": "http://192.168.1.164:8080/mcp",
        "sendHeaders": true,
        "headerParameters": {
          "parameters": [
            {
              "name": "Content-Type",
              "value": "application/json"
            }
          ]
        },
        "sendBody": true,
        "specifyBody": "json",
        "jsonBody": "{\n  \"jsonrpc\": \"2.0\",\n  \"id\": \"list-emails-folder\",\n  \"method\": \"list_emails\",\n  \"params\": {\n    \"folder_id\": \"00000000DB2820C5F3F8204492F273035529BA6801009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000\",\n    \"unread_only\": false,  \n    \"limit\": 10\n  }\n}",
        "options": {}
      },
      "type": "n8n-nodes-base.httpRequestTool",
      "typeVersion": 4.2,
      "position": [
        2432,
        -560
      ],
      "id": "07886ec6-6720-4d4c-9c03-5645a866a893",
      "name": "list emails"
    },
    {
      "parameters": {
        "toolDescription": "list email",
        "method": "POST",
        "url": "http://192.168.1.164:8080/mcp",
        "sendHeaders": true,
        "headerParameters": {
          "parameters": [
            {
              "name": "Content-Type",
              "value": "application/json"
            }
          ]
        },
        "sendBody": true,
        "specifyBody": "json",
        "jsonBody": "{\n  \"jsonrpc\": \"2.0\",\n  \"id\": \"list-inbox-emails\",\n  \"method\": \"list_inbox_emails\",\n  \"params\": {\n    \"limit\": 50\n  }\n}",
        "options": {}
      },
      "type": "n8n-nodes-base.httpRequestTool",
      "typeVersion": 4.2,
      "position": [
        2432,
        -384
      ],
      "id": "32d51ad9-8bd3-429d-a095-3779dd7cf489",
      "name": "list_inbox_emails"
    },
    {
      "parameters": {
        "toolDescription": "Fetches a list of the 100 most recent emails from the inbox. Returns a JSON array of email objects, each containing a subject, body, and received_time.",
        "method": "POST",
        "url": "http://192.168.1.164:8080/mcp",
        "sendHeaders": true,
        "headerParameters": {
          "parameters": [
            {
              "name": "Content-Type",
              "value": "application/json"
            }
          ]
        },
        "sendBody": true,
        "specifyBody": "json",
        "jsonBody": "{\n  \"jsonrpc\": \"2.0\",\n  \"id\": \"list-emails-folder\",\n  \"method\": \"list_inbox_emails\",\n  \"params\": {\n    \"limit\": 10\n  }\n}",
        "options": {}
      },
      "type": "n8n-nodes-base.httpRequestTool",
      "typeVersion": 4.2,
      "position": [
        1600,
        48
      ],
      "id": "8b0b1c17-c732-499f-b58d-f063bb244773",
      "name": "tool_list_emails",
      "notes": "Fetches a list of the 100 most recent emails from the inbox. Returns a JSON array of email objects, each containing a subject, body, and received_time."
    },
    {
      "parameters": {
        "modelName": "models/gemini-2.5-pro",
        "options": {}
      },
      "type": "@n8n/n8n-nodes-langchain.lmChatGoogleGemini",
      "typeVersion": 1,
      "position": [
        1328,
        48
      ],
      "id": "414add66-1fef-42df-ac49-d0ee248dc8ec",
      "name": "Google Gemini Chat Model",
      "credentials": {
        "googlePalmApi": {
          "id": "U6Pcb79nIRa1KZD2",
          "name": "Google Gemini(PaLM) Api account"
        }
      }
    },
    {
      "parameters": {
        "promptType": "define",
        "text": "You are a data extraction and JSON formatting agent. Your sole purpose is to analyze a list of emails, extract key information, and format it as a clean JSON array.\n\n**Analysis Rules:**\n1.  Call the `tool_list_emails` to get the email data.\n2.  Analyze the results and find all emails that match these criteria: the `subject` contains 'blade' OR the `body` contains 'SA', 'software team', 'driver', or 'image'.\n\n**Output Formatting Rules:**\n1.  For the matching emails you found, create a JSON array of objects.\n2.  Each object in the array MUST have these exact keys and value types:\n    - `emailTitle` (string)\n    - `mentionedKeywords` (string)\n    - `receivedTime` (string)\n    - `urgency` (string, e.g., \"High\", \"Normal\", \"Low\")\n3.  Your final output that you pass to the tool MUST be ONLY this JSON array, formatted as a single string. Do not include any other text, explanations, or markdown formatting.\n\n**Example of the required output format:**\n`[{\"emailTitle\":\"RE: [Blade] audio issue\",\"mentionedKeywords\":\"software team, SA\",\"receivedTime\":\"2025-09-23T17:29:51Z\",\"urgency\":\"High\"}]`\n\n**Execution:**\n- If you find matching emails, call the `tool_create_and_send_summary` tool and pass your generated JSON array string into the `query` parameter.\n- If you find no matching emails, respond with the exact text: \"No matching emails found.\"",
        "options": {}
      },
      "type": "@n8n/n8n-nodes-langchain.agent",
      "typeVersion": 2.2,
      "position": [
        1376,
        -160
      ],
      "id": "3c51d80a-f219-476d-8d78-74ed2ad039b4",
      "name": "AI Agent"
    },
    {
      "parameters": {
        "description": "Use this tool to generate and send the final email report. It takes one parameter, 'matchingEmails', which MUST be the JSON array of email objects that you decided were important.",
        "jsCode": "// The agent's output is directly available in a variable named 'query'.\n// This variable contains the string of summarized JSON data.\n\nlet summarizedEmails;\ntry {\n  // Parse the clean JSON string from the agent into a usable array.\n  summarizedEmails = JSON.parse(query);\n} catch (error) {\n  console.error(\"Failed to parse summary JSON from agent:\", error, \"Raw string was:\", query);\n  return \"Error: The summary data received from the agent was not valid JSON.\";\n}\n\nif (!summarizedEmails || summarizedEmails.length === 0) {\n  return \"Agent provided an empty summary. Nothing to send.\";\n}\n\n// --- 1. Generate HTML Table from the AI's Summary ---\nlet tableRows = '';\nfor (const summary of summarizedEmails) {\n  const title = summary.emailTitle || 'N/A';\n  const keywords = summary.mentionedKeywords || 'N/A';\n  const time = summary.receivedTime ? new Date(summary.receivedTime).toLocaleString() : 'N/A';\n  const urgency = summary.urgency || 'Normal';\n  \n  tableRows += `\n    <tr>\n      <td style=\"border: 1px solid #ddd; padding: 8px;\">${title}</td>\n      <td style=\"border: 1px solid #ddd; padding: 8px;\">${keywords}</td>\n      <td style=\"border: 1px solid #ddd; padding: 8px;\">${time}</td>\n      <td style=\"border: 1px solid #ddd; padding: 8px;\">${urgency}</td>\n    </tr>\n  `;\n}\n\nconst fullHtmlBody = `\n  <html><body><h2>Daily Email Summary</h2>\n  <table style=\"border-collapse: collapse; width: 100%; font-family: sans-serif;\">\n    <thead>\n      <tr>\n        <th style=\"border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2;\">Email Title</th>\n        <th style=\"border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2;\">Mentioned Keywords</th>\n        <th style=\"border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2;\">Received Time</th>\n        <th style=\"border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f2f2f2;\">Urgency</th>\n      </tr>\n    </thead>\n    <tbody>${tableRows}</tbody>\n  </table></body></html>`;\n\n// --- 2. Send the Email ---\nconst mcpPayload = {\n  jsonrpc: \"2.0\",\n  id: \"send-email-from-code-tool\",\n  method: \"send_email\",\n  params: {\n    to_recipients: [\"JackieCF_Lin@compal.com\"],\n    subject: `Daily Email Summary - ${new Date().toLocaleDateString()}`,\n    body: fullHtmlBody,\n    body_format: \"html\"\n  }\n};\n\ntry {\n  // --- THIS IS THE FIX ---\n  // Use n8n's built-in helper for making HTTP requests instead of 'fetch'.\n  await this.helpers.httpRequest({\n    method: 'POST',\n    url: 'http://192.168.1.164:8080/mcp',\n    headers: {\n      'Content-Type': 'application/json'\n    },\n    body: mcpPayload\n  });\n  return \"Success. The summary email was generated and sent.\";\n} catch (error) {\n  console.error('Failed to send email:', error);\n  return `Error: Failed to send email. Details: ${error.message}`;\n}"
      },
      "type": "@n8n/n8n-nodes-langchain.toolCode",
      "typeVersion": 1.3,
      "position": [
        1744,
        48
      ],
      "id": "6332ef1b-60d8-44bc-80bf-b4861890df30",
      "name": "tool_create_and_send_summary"
    },
    {
      "parameters": {
        "toolDescription": "**Function:** Sends a pre-formatted email summary.\n**Parameter:**\n* `query` (string, required): This MUST be a single string containing the complete, valid HTML for the body of the email.",
        "method": "POST",
        "url": "http://192.168.1.164:8080/mcp",
        "sendHeaders": true,
        "headerParameters": {
          "parameters": [
            {
              "name": "Content-Type",
              "value": "application/json"
            }
          ]
        },
        "sendBody": true,
        "contentType": "={\n  \"jsonrpc\": \"2.0\",\n  \"id\": \"send-email-working\",\n  \"method\": \"send_email\",\n  \"params\": {\n    \"to_recipients\": [\"JackieCF_Lin@compal.com\"],\n    \"subject\": \"Daily Email Summary - {{ new Date().toLocaleDateString() }}\",\n    \"body\": \"{{ $parameter.query}}\",\n    \"body_format\": \"html\"\n  }\n}",
        "bodyParameters": {
          "parameters": [
            {}
          ]
        },
        "options": {}
      },
      "type": "n8n-nodes-base.httpRequestTool",
      "typeVersion": 4.2,
      "position": [
        2144,
        -544
      ],
      "id": "3432dc48-54ca-4f67-be79-eda7ee1d631d",
      "name": "tool_send_email",
      "notes": "Sends an email with an HTML body to JackieCF_Lin@compal.com. This tool requires exactly one parameter: 'emailBody', which must contain the complete HTML table for the email content."
    }
  ],
  "connections": {
    "Code in JavaScript": {
      "main": [
        [
          {
            "node": "if '收件匣' in folders",
            "type": "main",
            "index": 0
          },
          {
            "node": "if 'inbox' in folders",
            "type": "main",
            "index": 0
          }
        ]
      ]
    },
    "MCP Server health check": {
      "main": [
        [
          {
            "node": "MCP Server alive",
            "type": "main",
            "index": 0
          }
        ]
      ]
    },
    "MCP Server alive": {
      "main": [
        [
          {
            "node": "POST /mcp /get_folders",
            "type": "main",
            "index": 0
          }
        ],
        [
          {
            "node": "Stop and Error",
            "type": "main",
            "index": 0
          }
        ]
      ]
    },
    "POST /mcp /get_folders": {
      "main": [
        [
          {
            "node": "Code in JavaScript",
            "type": "main",
            "index": 0
          }
        ]
      ]
    },
    "Code in Python (Beta)": {
      "main": [
        [
          {
            "node": "AI Agent",
            "type": "main",
            "index": 0
          }
        ]
      ]
    },
    "if 'inbox' in folders": {
      "main": [
        [
          {
            "node": "Code in Python (Beta)1",
            "type": "main",
            "index": 0
          }
        ]
      ]
    },
    "if '收件匣' in folders": {
      "main": [
        [
          {
            "node": "Code in Python (Beta)",
            "type": "main",
            "index": 0
          }
        ],
        []
      ]
    },
    "Get unread emails": {
      "ai_tool": [
        [
          {
            "node": "Message a model",
            "type": "ai_tool",
            "index": 0
          }
        ]
      ]
    },
    "Trigger at 8am everyday": {
      "main": [
        [
          {
            "node": "MCP Server health check",
            "type": "main",
            "index": 0
          }
        ]
      ]
    },
    "search_emails": {
      "ai_tool": [
        [
          {
            "node": "Message a model",
            "type": "ai_tool",
            "index": 0
          }
        ]
      ]
    },
    "send_mails": {
      "ai_tool": [
        [
          {
            "node": "Message a model",
            "type": "ai_tool",
            "index": 0
          }
        ]
      ]
    },
    "get email(id)": {
      "ai_tool": [
        [
          {
            "node": "Message a model",
            "type": "ai_tool",
            "index": 0
          }
        ]
      ]
    },
    "list emails": {
      "ai_tool": [
        [
          {
            "node": "Message a model",
            "type": "ai_tool",
            "index": 0
          }
        ]
      ]
    },
    "list_inbox_emails": {
      "ai_tool": [
        []
      ]
    },
    "tool_list_emails": {
      "ai_tool": [
        [
          {
            "node": "AI Agent",
            "type": "ai_tool",
            "index": 0
          }
        ]
      ]
    },
    "Google Gemini Chat Model": {
      "ai_languageModel": [
        [
          {
            "node": "AI Agent",
            "type": "ai_languageModel",
            "index": 0
          }
        ]
      ]
    },
    "AI Agent": {
      "main": [
        []
      ]
    },
    "tool_create_and_send_summary": {
      "ai_tool": [
        [
          {
            "node": "AI Agent",
            "type": "ai_tool",
            "index": 0
          }
        ]
      ]
    },
    "tool_send_email": {
      "ai_tool": [
        []
      ]
    }
  },
  "pinData": {},
  "meta": {
    "templateCredsSetupCompleted": true,
    "instanceId": "9332abf076bbc6feac3987406a626794a341818b9eab3b4148a8b2fe16d5a1dd"
  }
}