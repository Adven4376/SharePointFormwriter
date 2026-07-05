/**
 * office.js - Mock Office.js library.
 * This is loaded locally to support developer testing of the Task Pane UI in standard web browsers.
 * Inside Outlook, the real Office.js is loaded from the Microsoft CDN.
 */
if (typeof Office === 'undefined') {
    window.Office = {
        initialize: function () {},
        onReady: function (callback) {
            console.log("[Mock Office.js] Mock initialization ready event triggered.");
            // Simulate asynchronous loading callback
            setTimeout(() => {
                callback({
                    host: window.Office.HostType.Outlook,
                    platform: window.Office.PlatformType.PC
                });
            }, 200);
        },
        HostType: {
            Outlook: "Outlook",
            Word: "Word",
            Excel: "Excel",
            PowerPoint: "PowerPoint"
        },
        PlatformType: {
            PC: "PC",
            Mac: "Mac",
            OfficeOnline: "OfficeOnline",
            iOS: "iOS",
            Android: "Android"
        },
        context: {
            mailbox: {
                userProfile: {
                    displayName: "Aditya Developer",
                    emailAddress: "aditya.dev@enterprise.com"
                }
            }
        }
    };
    
    // Add warning styling to the header when running in mock mode
    window.addEventListener("DOMContentLoaded", () => {
        const header = document.querySelector(".app-header");
        if (header) {
            const badge = document.createElement("div");
            badge.style.cssText = "background-color: hsla(38, 95%, 50%, 0.15); border: 1px solid hsla(38, 95%, 50%, 0.3); color: hsl(38, 95%, 60%); font-size: 10px; padding: 4px 8px; border-radius: 4px; text-align: center; margin-top: 8px; font-weight: 500; font-family: sans-serif;";
            badge.innerText = "⚠️ STANDALONE WEB PREVIEW MODE (Mock Office.js)";
            header.appendChild(badge);
        }
    });
}
