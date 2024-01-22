Office.onReady(function (info) {
    // Office is ready
    if (info.host === Office.HostType.Outlook) {
        // Assign event handler to itemRead event
        Office.item.addHandlerAsync(Office.EventType.ItemRead, onItemRead);
    }
});

// Event handler for itemChanged event
async function onItemRead(eventArgs) {
    try {
        // Get current item
        const currentItem = Office.context.mailbox.item;

        // Check if the item has attachments
        if (currentItem.attachments.length > 0) {
            // Iterate through attachments
            for (const attachment of currentItem.attachments) {
                // Check if the attachment is an image
                if (attachment.contentType.includes('image') || attachment.contentType.includes('pdf')) {
                    // Call OpenAI GPT-4 Vision API for signature detection
                    const attachmentId = `${attachment.id}`;
                    const hasSignature = await checkForSignature(attachmentId);

                    // Display whether each file has a signature
                    displaySignatureResult(attachment.name, hasSignature);

                }
            }
        }
    } catch (error) {
        console.error('Error:', error);
    }
}

// Function to call OpenAI GPT-4 Vision API for signature detection
async function checkForSignature(attachmentId) {
    // Replace 'YOUR_OPENAI_API_KEY' with your actual OpenAI API key
    const openaiApiKey = 'YOUR_OPENAI_API_KEY';

    // Retrieve attachment data as a buffer
    const imageBuffer = await Office.context.mailbox.item.getAttachmentContentAsync(attachmentId);

    // Convert buffer to Base64-encoded string
    const base64Image = imageBuffer.toString('base64');

    try {
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${openaiApiKey}`,
            },
            body: JSON.stringify({
                model: 'gpt-4-vision-preview',
                messages: [
                    {
                        role: 'user',
                        content: `Does this image contain a signature? Respond with either 'yes' or 'no' and nothing else: ${base64Image}`,
                    },
                ],
                max_tokens: 150,
                n: 1,
                stop: null,
                temperature: 0.5,
            }),
        });

        const responseBody = await response.json();
        const answer = responseBody.choices[0].message.content;
        return answer.includes('yes');

    } catch (error) {
        console.error('Error checking for signature:', error);
        return false;
    }
}

// Function to display the result in the body of the email
function displaySignatureResult(attachmentName, hasSignature) {
    const body = Office.context.mailbox.item.body;
    const signatureStatus = hasSignature ? 'Signature Detected' : 'No Signature Detected';

    // Add HTML to the body of the email
    body.setAsync(
        `<p>${attachmentName}: ${signatureStatus}</p>`,
        { coercionType: Office.CoercionType.Html }
    );
}
