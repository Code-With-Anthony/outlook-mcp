const graphClient = require('./graphClient');

/**
 * Fetches the most recent emails from the inbox.
 */
async function fetchEmails(top = 10) {
    try {
        const result = await graphClient
            .api('/me/messages')
            .select('subject,sender,bodyPreview,receivedDateTime')
            .orderby('receivedDateTime DESC')
            .top(top)
            .get();
            
        return result.value.map(message => ({
            subject: message.subject,
            sender: message.sender?.emailAddress?.name || message.sender?.emailAddress?.address,
            senderEmail: message.sender?.emailAddress?.address,
            preview: message.bodyPreview,
            received: message.receivedDateTime
        }));
    } catch (error) {
        throw new Error(`Failed to fetch emails: ${error.message}`);
    }
}

/**
 * Searches emails based on a keyword.
 */
async function searchEmails(keyword) {
    try {
        const result = await graphClient
            .api('/me/messages')
            // Using standard Graph $search
            .search(`"${keyword}"`)
            .select('subject,sender,bodyPreview,receivedDateTime')
            .top(15)
            .get();
            
        return result.value.map(message => ({
            subject: message.subject,
            sender: message.sender?.emailAddress?.name || message.sender?.emailAddress?.address,
            preview: message.bodyPreview,
            received: message.receivedDateTime
        }));
    } catch (error) {
        throw new Error(`Failed to search emails: ${error.message}`);
    }
}

/**
 * Sends a plain-text email.
 */
async function sendEmail(to, subject, body) {
    try {
        const message = {
            message: {
                subject: subject,
                body: {
                    contentType: "Text",
                    content: body
                },
                toRecipients: [
                    {
                        emailAddress: {
                            address: to
                        }
                    }
                ]
            },
            saveToSentItems: "true" // Keep a trace in sent box
        };

        await graphClient.api('/me/sendMail').post(message);
        return { status: "success", message: `Email successfully sent to ${to}` };
    } catch (error) {
        throw new Error(`Failed to send email: ${error.message}`);
    }
}

/**
 * Retrieves upcoming calendar events for the specified number of days.
 */
async function getCalendarEvents(daysLookAhead = 7) {
    try {
        const now = new Date();
        const future = new Date();
        future.setDate(now.getDate() + daysLookAhead);
        
        // Use calendarView to expand recurring events properly
        const startDateTime = now.toISOString();
        const endDateTime = future.toISOString();
        
        const result = await graphClient
            .api(`/me/calendarView?startDateTime=${startDateTime}&endDateTime=${endDateTime}`)
            .select('subject,start,end,location,organizer')
            .orderby('start/dateTime ASC')
            .top(20)
            .get();
            
        return result.value.map(event => ({
            subject: event.subject,
            start: event.start?.dateTime,
            end: event.end?.dateTime,
            timeZone: event.start?.timeZone,
            location: event.location?.displayName || "No location specified",
            organizer: event.organizer?.emailAddress?.name
        }));
    } catch (error) {
        throw new Error(`Failed to get calendar events: ${error.message}`);
    }
}

/**
 * Creates a new calendar event.
 */
async function createEvent(title, startTime, endTime) {
    try {
        const event = {
            subject: title,
            start: {
                dateTime: startTime,
                timeZone: "UTC" // AI models deal with UTC easily
            },
            end: {
                dateTime: endTime,
                timeZone: "UTC"
            }
        };

        const result = await graphClient.api('/me/events').post(event);
        return { 
            status: "success", 
            message: "Event created successfully.",
            eventId: result.id,
            webLink: result.webLink 
        };
    } catch (error) {
        throw new Error(`Failed to create event: ${error.message}`);
    }
}

module.exports = {
    fetchEmails,
    searchEmails,
    sendEmail,
    getCalendarEvents,
    createEvent
};
