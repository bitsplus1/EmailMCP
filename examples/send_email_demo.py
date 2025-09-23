"""Demo script showing how to use the send_email functionality."""

import json
import asyncio
from src.outlook_mcp_server.server import OutlookMCPServer


async def demo_send_email():
    """Demonstrate send_email functionality."""
    
    print("📧 Outlook MCP Server - Send Email Demo")
    print("=" * 50)
    
    # Create and start server
    server = OutlookMCPServer()
    
    try:
        print("🚀 Starting MCP server...")
        await server.start()
        print("✅ Server started successfully\n")
        
        # Example 1: Simple text email
        print("📝 Example 1: Simple Text Email")
        simple_request = {
            "jsonrpc": "2.0",
            "id": "simple-email",
            "method": "send_email",
            "params": {
                "to_recipients": ["recipient@example.com"],
                "subject": "Simple Test Email",
                "body": "Hello! This is a simple text email sent via the MCP server.",
                "body_format": "text"
            }
        }
        
        print("Request:")
        print(json.dumps(simple_request, indent=2))
        print("\n" + "─" * 50 + "\n")
        
        # Example 2: HTML email with CC and attachments
        print("📝 Example 2: Rich HTML Email with CC and Attachments")
        rich_request = {
            "jsonrpc": "2.0",
            "id": "rich-email",
            "method": "send_email",
            "params": {
                "to_recipients": ["primary@example.com", "secondary@example.com"],
                "cc_recipients": ["manager@example.com"],
                "bcc_recipients": ["archive@example.com"],
                "subject": "📊 Weekly Report - Automated",
                "body": """
                <html>
                <body>
                    <h1 style="color: #2E86AB;">Weekly Report</h1>
                    <p>Dear Team,</p>
                    <p>Please find the weekly report attached. Key highlights:</p>
                    <ul>
                        <li><strong>Tasks Completed:</strong> 25</li>
                        <li><strong>Issues Resolved:</strong> 8</li>
                        <li><strong>New Features:</strong> 3</li>
                    </ul>
                    <p>Best regards,<br>
                    <em>Automated Reporting System</em></p>
                </body>
                </html>
                """,
                "body_format": "html",
                "importance": "high",
                "attachments": [
                    "C:\\reports\\weekly_summary.pdf",
                    "C:\\reports\\metrics.xlsx"
                ]
            }
        }
        
        print("Request:")
        print(json.dumps(rich_request, indent=2))
        print("\n" + "─" * 50 + "\n")
        
        # Example 3: Notification email
        print("📝 Example 3: System Notification Email")
        notification_request = {
            "jsonrpc": "2.0",
            "id": "notification-email",
            "method": "send_email",
            "params": {
                "to_recipients": ["admin@example.com"],
                "subject": "🚨 System Alert: High CPU Usage Detected",
                "body": """
                <html>
                <body style="font-family: Arial, sans-serif;">
                    <div style="background-color: #f8d7da; border: 1px solid #f5c6cb; padding: 15px; border-radius: 5px;">
                        <h2 style="color: #721c24; margin-top: 0;">⚠️ System Alert</h2>
                        <p><strong>Alert Type:</strong> High CPU Usage</p>
                        <p><strong>Server:</strong> web-server-01</p>
                        <p><strong>Current Usage:</strong> 95%</p>
                        <p><strong>Threshold:</strong> 80%</p>
                        <p><strong>Time:</strong> 2025-09-22 14:30:00 UTC</p>
                    </div>
                    <p>Please investigate immediately.</p>
                    <p><em>This is an automated alert from the monitoring system.</em></p>
                </body>
                </html>
                """,
                "body_format": "html",
                "importance": "high"
            }
        }
        
        print("Request:")
        print(json.dumps(notification_request, indent=2))
        print("\n" + "─" * 50 + "\n")
        
        # Show usage in n8n
        print("🔧 Usage in n8n HTTP Request Node:")
        print("1. Method: POST")
        print("2. URL: http://127.0.0.1:8080/mcp")
        print("3. Headers: Content-Type: application/json")
        print("4. Body: Use any of the JSON examples above")
        print("\n" + "─" * 50 + "\n")
        
        # Show common use cases
        print("💡 Common Use Cases:")
        use_cases = [
            "📊 Automated Reports: Send daily/weekly reports with data attachments",
            "🚨 System Alerts: Notify administrators of system issues or thresholds",
            "📋 Task Notifications: Update team members on task completions or assignments",
            "📈 Data Export: Email processed data or analysis results to stakeholders",
            "🔄 Workflow Updates: Send status updates for automated workflows",
            "📝 Form Submissions: Email form data to relevant departments",
            "⏰ Scheduled Reminders: Send automated reminders for meetings or deadlines",
            "🎯 Marketing Automation: Send personalized emails based on triggers"
        ]
        
        for use_case in use_cases:
            print(f"  • {use_case}")
        
        print("\n" + "─" * 50 + "\n")
        
        print("✨ Features:")
        features = [
            "✅ Multiple recipients (TO, CC, BCC)",
            "✅ HTML, Text, and RTF body formats",
            "✅ File attachments support",
            "✅ Importance levels (low, normal, high)",
            "✅ Email validation and error handling",
            "✅ Rate limiting and performance optimization",
            "✅ Comprehensive logging and monitoring"
        ]
        
        for feature in features:
            print(f"  {feature}")
        
        print(f"\n🎉 send_email function is ready to use!")
        print(f"📚 See docs/N8N_INTEGRATION_SETUP.md for complete integration guide")
        
    except Exception as e:
        print(f"❌ Demo failed: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        try:
            await server.stop()
            print("\n🛑 Server stopped")
        except:
            pass


if __name__ == "__main__":
    asyncio.run(demo_send_email())