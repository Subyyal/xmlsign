/* global Office */

Office.onReady(() => {
    // Office is initialized
});

/**
 * Event-based handler triggered whenever a user composes a new message or reply
 */
function onNewMessageComposeHandler(event) {
    insertCompanySignature(event, false);
}

function onNewAppointmentComposeHandler(event) {
    insertCompanySignature(event, false);
}

/**
 * Inserts the signature into Outlook compose window
 * Uses Office.context.mailbox.item.body.setSignatureAsync
 */
function insertCompanySignature(event, isReply) {
    try {
        const userProfile = Office.context.mailbox.userProfile;
        
        // Dynamically populate details from the logged-in Outlook user profile
        const displayName = userProfile.displayName || "Ahmed Al-Mansoor";
        const emailAddress = userProfile.emailAddress || "ahmed.m@company.com";
        const jobTitle = "Senior Structural Engineer";
        const department = "Engineering Operations";
        const company = "Enterprise Engineering Consultants";
        const mobile = "+971 50 123 4567";
        const office = "+971 4 987 6543";
        const website = "www.company.com";
        const logoUrl = "https://img.icons8.com/color/96/briefcase.png";

        // Compact HTML Table strictly formatted for Outlook (Zero gaps, fixed width, responsive)
        const mobileRow = mobile ? `
            <tr>
                <td style="font-family: Arial, Helvetica, sans-serif; font-size: 12px; line-height: 16px; color: #555555; padding-top: 2px;">
                    <span style="color: #0078d4; font-weight: bold;">M:</span> 
                    <a href="tel:${mobile}" style="color: #555555; text-decoration: none;">${mobile}</a>
                </td>
            </tr>` : "";

        const signatureHtml = `
            <div data-signature-id="company-managed-signature" style="font-family: Arial, Helvetica, sans-serif; margin-top: 20px;">
                <table cellpadding="0" cellspacing="0" border="0" style="width: 100%; max-width: 480px; font-family: Arial, Helvetica, sans-serif; border-collapse: collapse;">
                    <tr>
                        <!-- Company Logo -->
                        <td valign="top" style="width: 65px; padding-right: 15px; border-right: 2px solid #0078d4;">
                            <img src="${logoUrl}" width="55" height="55" alt="Logo" style="display: block; width: 55px; height: 55px; border: 0;" />
                        </td>
                        
                        <!-- Employee Info -->
                        <td valign="top" style="padding-left: 15px;">
                            <table cellpadding="0" cellspacing="0" border="0" style="border-collapse: collapse;">
                                <tr>
                                    <td style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 18px; font-weight: bold; color: #111111;">
                                        ${displayName}
                                    </td>
                                </tr>
                                <tr>
                                    <td style="font-family: Arial, Helvetica, sans-serif; font-size: 12px; line-height: 16px; color: #0078d4; font-weight: 600;">
                                        ${jobTitle} &nbsp;|&nbsp; ${department}
                                    </td>
                                </tr>
                                <tr>
                                    <td style="font-family: Arial, Helvetica, sans-serif; font-size: 12px; line-height: 16px; color: #666666; font-weight: bold; padding-top: 2px;">
                                        ${company}
                                    </td>
                                </tr>
                                ${mobileRow}
                                <tr>
                                    <td style="font-family: Arial, Helvetica, sans-serif; font-size: 12px; line-height: 16px; color: #555555; padding-top: 2px;">
                                        <span style="color: #0078d4; font-weight: bold;">T:</span> ${office} &nbsp;|&nbsp; 
                                        <a href="mailto:${emailAddress}" style="color: #0078d4; text-decoration: none;">${emailAddress}</a>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="font-family: Arial, Helvetica, sans-serif; font-size: 11px; line-height: 15px; color: #888888; padding-top: 6px;">
                                        <a href="https://${website}" target="_blank" style="color: #0078d4; text-decoration: none; font-weight: bold;">${website}</a> &nbsp;|&nbsp; 
                                        <em>CONFIDENTIALITY NOTICE: This transmission is intended only for the designated recipient.</em>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </div>
        `;

        // Check if setSignatureAsync is supported (Mailbox requirement set 1.10+)
        if (Office.context.mailbox.item.body.setSignatureAsync) {
            Office.context.mailbox.item.body.setSignatureAsync(
                signatureHtml,
                { coercionType: Office.CoercionType.Html },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error("Failed to set signature: " + asyncResult.error.message);
                    }
                    event.completed();
                }
            );
        } else {
            // Fallback for older Office API versions: append to body
            Office.context.mailbox.item.body.setSelectedDataAsync(
                signatureHtml,
                { coercionType: Office.CoercionType.Html },
                function () {
                    event.completed();
                }
            );
        }
    } catch (err) {
        console.error("Error setting signature: ", err);
        event.completed();
    }
}
