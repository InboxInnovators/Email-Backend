import express from 'express';
import cors from 'cors';
import axios from 'axios';
import 'dotenv/config';
import path from 'path';
import fs from 'fs';
import { JSDOM } from 'jsdom';

import { fileURLToPath } from 'url';
import { dirname } from 'path';
import { GoogleGenerativeAI } from '@google/generative-ai';
import * as jsforce from 'jsforce'; // Import jsforce for Salesforce connection
import { ChatGoogleGenerativeAI } from '@langchain/google-genai';
import { ChatPromptTemplate } from "@langchain/core/prompts";
import { StringOutputParser } from "@langchain/core/output_parsers";
import http from 'http';
import WebSocket from 'ws';


const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const app = express();
app.use(cors());

// Initialize Google Generative AI
// Initialize Gemini with 1.5 Flash model
const genAI = new GoogleGenerativeAI(process.env.GOOGLE_API_KEY);
const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

// Middleware
app.use(cors({
  origin: process.env.FRONTEND_URL, // Your React app's URL
}));
app.use(express.json()); // Middleware to parse JSON bodies

// Endpoint to fetch emails using the access token
app.post('/api/emails', async (req, res) => {
    const { accessToken } = req.body; // Get access token from request body

    if (!accessToken) {
        return res.status(401).json({ message: 'Access token is required' });
    }

    try {
        // Fetch user's emails using Graph API
        const response = await axios.get("https://graph.microsoft.com/v1.0/me/messages?$top=50", {
            headers: {
                Authorization: `Bearer ${accessToken}`,
            },
        });

        const emails = response.data.value; // Extract the emails
        console.log(emails)
        // Fetch folders to get parent folder names
        const foldersResponse = await axios.get('https://graph.microsoft.com/v1.0/me/mailFolders', {
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });

        // Create a mapping of folder ID to folder name
        const folderMap = {};
        foldersResponse.data.value.forEach(folder => {
            folderMap[folder.id] = folder.displayName; // Map folder ID to display name
        });

        // Add parent folder name to each email
        const emailsWithFolderNames = emails.map(email => ({
            ...email,
            parentFolderName: folderMap[email.parentFolderId] || 'Unknown Folder' // Add parent folder name
        }));

        res.status(200).json({ emails: emailsWithFolderNames });
    } catch (error) {
        console.error("Error fetching emails:", error);
        return res.status(500).json({ message: "Error fetching emails" });
    }
});

// Endpoint to get folders
app.post("/api/folders", async (req, res) => {
    const { accessToken } = req.body; // Get access token from request body

    if (!accessToken) {
        return res.status(401).json({ message: 'Access token is required' });
    }

    try {
        const response = await axios.get(
            'https://graph.microsoft.com/v1.0/me/mailFolders',
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`
                }
            }
        );

        // Log the fetched folders for debugging
       // console.log('Fetched folders:', response.data);

        // Return the response data in the desired format
        return res.status(200).json({
            '@odata.context': response.data['@odata.context'],
            value: response.data.value.map(folder => ({
                id: folder.id,
                displayName: folder.displayName,
                parentFolderId: folder.parentFolderId,
                childFolderCount: folder.childFolderCount,
                unreadItemCount: folder.unreadItemCount,
                totalItemCount: folder.totalItemCount,
                sizeInBytes: folder.sizeInBytes,
                isHidden: folder.isHidden
            }))
        });
    } catch (error) {
        //console.error('Error fetching folders:', error);
        return res.status(500).json({ error: 'Failed to fetch folders' });
    }
});

// //Endpoint to summarize the email using AI
// app.post('/api/summarize', async (req, res) => {
//     try {
//         const { body } = req.body;

//         // Validate email content
//         if (!body || typeof body !== 'string') {
//             return res.status(400).json({ error: 'Invalid email content provided' });
//         }

//         // Construct the prompt
//         const prompt = `Summarize the following email in the most concise way possible, focusing on the key points such as the purpose of the email, any actions required, deadlines, or important details. Adjust the length of the summary based on the email's content; if the email is short, provide a proportionately brief summary. `;
//         const userInput = body;

//         // Generate summary using Google Generative AI
//         const result = await model.generateContent(prompt + userInput);

//         // Send the summary back to the client
//         const summary = result.response.text(); // Ensure this is correct
//         return res.status(200).json({ summary });
//     } catch (error) {
//         console.error('Error summarizing email:', error);
//         return res.status(500).json({ error: 'Failed to summarize email' });
//     }
// });

// Endpoint summarize stream 

// // Endpoint to summarize the email using AI
// app.post('/api/summarize', async (req, res) => {
//     try {
//         const { body } = req.body;

//         // Validate email content
//         if (!body || typeof body !== 'string') {
//             return res.status(400).json({ error: 'Invalid email content provided' });
//         }
//         // masked body
//         // Mask the PII
//         const maskedText = maskPII(body);
//         const demaskedText=demaskPII(maskedText);
//         console.log('Masked Text::::', maskedText);
//         console.log('Demasked Text:::',demaskedText);

//         // Construct the prompt
//         const prompt = `Summarize the following email in the most concise way possible, focusing on the key points such as the purpose of the email, any actions required, deadlines, or important details. Adjust the length of the summary based on the email's content; if the email is short, provide a proportionately brief summary. `;
//         const userInput = body;

//         // Use the Gemini AI streaming method to generate summary
//         const result = await model.generateContentStream(prompt + maskedText);

//         // Set headers for streaming
//         res.setHeader('Content-Type', 'text-event-stream');
//         res.setHeader('Cache-Control', 'no-cache');
//         res.setHeader('Connection', 'keep-alive');

//         // Stream the response
//         for await (const chunk of result.stream) {
//             const chunkText = chunk.text(); // Extract the text from the chunk
//             res.write(`data: ${chunkText}\n\n`); // Send the chunk to the client
//         }

//         res.end(); // End the response after streaming is complete

//     } catch (error) {
//         console.error('Error summarizing email:', error);
//         return res.status(500).json({ error: 'Failed to summarize email' });
//     }
// });
let llm=null;

function getLlmInstance() {
    if (!llm) {
        llm = new ChatGoogleGenerativeAI({
            model: "gemini-1.5-flash",
            stream: true, // Enable streaming
            
        });
    }
    return llm;
}

app.post('/api/summarize', async (req, res) => {
    try {
        const { body } = req.body;

        // Validate the request
        if (!body || typeof body !== 'string') {
            return res.status(400).json({ error: 'Invalid email content provided' });
        }

        const emailToAnalyze = body.trim(); // Replace with your `maskPII` logic if needed

        // Construct the prompt for AI
        const prompt = `Summarize the following email in the most concise way possible, focusing on key points such as purpose, actions required, deadlines, or important details:\n\nEmail: ${emailToAnalyze}`;

        

        // Generate the summary
        const result = await model.generateContent(prompt); // Use generateContent to generate the summary

        // Extract the generated summary content
        const summary = result.response.text(); // Ensure this matches your model's response structure

        // Send the summary back to the client
        res.status(200).json({ summary });
    } catch (error) {
        console.error('Error summarizing email:', error);
        res.status(500).json({ error: 'Failed to summarize email' });
    }
});
// New endpoint to send and save the email with CC
app.post('/api/sendEmail', async (req, res) => {
    const { accessToken, subject, body, recipients, ccRecipients } = req.body; // Get access token and email details from request body

    if (!accessToken) {
        return res.status(401).json({ message: 'Access token is required' });
    }

    if (!subject || !body || !recipients || !Array.isArray(recipients)) {
        return res.status(400).json({ message: 'Subject, body, and recipients are required' });
    }

    // // Format recipients for the API
    // const formattedRecipients = recipients.map(email => ({
    //     emailAddress: {
    //         address: email
    //     }
    // }));

    const message = {
        subject: subject,
        body: {
            contentType: "Text",
            content: body
        },
        toRecipients: recipients.map(email => ({
            emailAddress: {
                address: email
            }
        })),
        // Add ccRecipients if provided
        ccRecipients: ccRecipients && Array.isArray(ccRecipients)
            ? ccRecipients.map(email => ({
                emailAddress: {
                    address: email
                }
            }))
            : []
    };

    try {
        const response = await axios.post(
            'https://graph.microsoft.com/v1.0/me/sendMail',
            {
                message: message
            },
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            }
        );
        console.log('Email sent successfully:', response.data);
        return res.status(200).json({ message: 'Email sent successfully' });
    } catch (error) {
        console.error('Error sending email:', error.response ? error.response.data : error.message);
        return res.status(500).json({ message: 'Error sending email' });
    }
});


//Language translation endpoint 
app.post('/translate', async (req, res) => {
    const { text, sourceLanguage, targetLanguage } = req.body;

    console.log('Text:', text);
    console.log('Source Language:', sourceLanguage);
    console.log('Target Language:', targetLanguage);

    // Validate input
    if (!text || !sourceLanguage || !targetLanguage) {
        return res.status(400).json({ error: 'Text, sourceLanguage, and targetLanguage are required' });
    }

    try {
        // Ensure the prompt is formatted correctly
        const response = await model.generateContent(`You are a Translate API model, translate the following html email body from ${sourceLanguage} to ${targetLanguage}: ${text} in plain text format. don't give any other warnings or explaination`);
        const result=response.response.text();
        console.log('Response:',result)

        res.json({ result });
    } catch (error) {
        console.error('Error during translation:', error);
        res.status(500).json({ error: 'Translation failed' });
    }
});

// New endpoint to get email details from Salesforce
app.post('/api/getEmailDetails', async (req, res) => {
    const { email } = req.body; // Get email from request body

    if (!email) {
        return res.status(400).json({ message: 'Email is required' });
    }

    const username = 'joiningapp7@resourceful-fox-7w3tc4.com'; // Salesforce username
    const password = 'Hello@123123CwozoRmQ28h4Fkx24mCyNU8q'; // Salesforce password
    const connection = new jsforce.Connection({
        loginUrl: 'https://login.salesforce.com'
    });

    try {
        await connection.login(username, password); // Log in to Salesforce
        console.log('Logged in successfully!');

        const details = await getEmailDetails(connection, email); // Get email details
        res.status(200).json({ details }); // Send details back to client
    } catch (error) {
        console.error('Error during Salesforce operation:', error);
        return res.status(500).json({ message: 'Error fetching email details from Salesforce' });
    }
});

// Function to get email details from Salesforce
async function getEmailDetails(connection, email) {
    const endpoint = `/EmailSearchSummary?email=${email}`;
    try {
        const response = await connection.apex.get(endpoint);
        console.log('Response:', response);
        return response;
    } catch (e) {
        console.error('Error during REST call:', e);
        throw e; // Rethrow the error to handle it in the endpoint
    }
}

// New endpoint to delete a message/email
app.delete('/api/deleteMessage', async (req, res) => {
    const { accessToken, messageId } = req.body; // Get access token and message ID from request body

    if (!accessToken || !messageId) {
        return res.status(400).json({ message: 'Access token and message ID are required' });
    }

    try {
        await deleteMessage(accessToken, messageId); // Call the deleteMessage function
        return res.status(200).json({ message: 'Email deleted successfully' });
    } catch (error) {
        console.error('Error deleting email:', error);
        return res.status(500).json({ message: 'Error deleting email' });
    }
});

// Function to delete a message/email
async function deleteMessage(token, messageId) {
    const encodedId = encodeURIComponent(messageId);
    const response = await axios.delete(
        `https://graph.microsoft.com/v1.0/me/messages/${encodedId}`,
        {
            headers: {
                Authorization: `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        }
    );
    console.log('The email was deleted');
    console.dir(response.data);
}

// New endpoint to fetch attachments for a specific email
app.get('/api/attachments/:emailId', async (req, res) => {
    const { emailId } = req.params; // Get email ID from request parameters
    const { accessToken } = req.body; // Get access token from request body

    if (!accessToken) {
        return res.status(401).json({ message: 'Access token is required' });
    }

    try {
        const attachments = await fetchAttachments(emailId, accessToken); // Fetch attachments
        return res.status(200).json({ attachments }); // Send attachments back to client
    } catch (error) {
       // console.error('Error fetching attachments:', error);
        return res.status(500).json({ message: 'Error fetching attachments' });
    }
});

// Function to fetch attachments for a specific email
const fetchAttachments = async (emailId, accessToken) => {
    try {
        const response = await fetch(
            `https://graph.microsoft.com/v1.0/me/messages/${emailId}/attachments`,
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`, // Pass your access token
                },
            }
        );

        if (!response.ok) {
            throw new Error('Failed to fetch attachments');
        }

        const data = await response.json();
        return data.value; // Array of attachment objects
    } catch (error) {
       // console.error('Error fetching attachments:', error);
        return [];
    }
};

// Create a new folder
app.post('/api/createFolder', async (req, res) => {
    const { accessToken, folderName } = req.body; // Get access token and folder name from request body

    if (!accessToken || !folderName) {
        return res.status(400).json({ message: 'Access token and folder name are required' });
    }

    try {
        const parentFolderId = 'msgfolderroot'; // Use fixed parentFolderId
        await createFolder(accessToken, parentFolderId, folderName); // Call the createFolder function
        return res.status(201).json({ message: 'Folder created successfully' });
    } catch (error) {
        return res.status(500).json({ message: 'Error creating folder' });
    }
});

// Rename a folder
app.patch('/api/renameFolder', async (req, res) => {
    const { accessToken, folderId, newFolderName } = req.body; // Get access token, folder ID, and new folder name from request body

    if (!accessToken || !folderId || !newFolderName) {
        return res.status(400).json({ message: 'Access token, folder ID, and new folder name are required' });
    }

    try {
        await renameFolder(accessToken, folderId, newFolderName); // Call the renameFolder function
        return res.status(200).json({ message: 'Folder renamed successfully' });
    } catch (error) {
        return res.status(500).json({ message: 'Error renaming folder' });
    }
});

// Delete a folder by name
app.delete('/api/deleteFolder', async (req, res) => {
    const { accessToken, folderName } = req.body; // Get access token and folder name from request body
    if (!accessToken || !folderName) {
        return res.status(400).json({ message: 'Access token and folder name are required' });
    }

    try {
        await deleteFolderByName(folderName, accessToken); // Call the deleteFolderByName function
        return res.status(200).json({ message: `Folder "${folderName}" deleted successfully` });
    } catch (error) {
        return res.status(500).json({ message: `Error deleting folder: ${error.message}` });
    }
});

// Function to delete a folder by name
async function deleteFolderByName(folderName, accessToken) {
    try {
        // Step 1: Fetch all folders
        const response = await fetch("https://graph.microsoft.com/v1.0/me/mailFolders", {
            method: "GET",
            headers: {
                "Authorization": `Bearer ${accessToken}`,
                "Content-Type": "application/json"
            }
        });
        
        if (!response.ok) throw new Error("Failed to fetch mail folders");

        const data = await response.json();
        const folder = data.value.find(f => f.displayName === folderName);

        if (!folder) {
            console.log(`Folder with name "${folderName}" not found.`);
            throw new Error(`Folder with name "${folderName}" not found.`);
        }

        // Step 2: Delete the folder
        const deleteResponse = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${folder.id}`, {
            method: "DELETE",
            headers: {
                "Authorization": `Bearer ${accessToken}`
            }
        });

        if (deleteResponse.ok) {
            console.log(`Folder "${folderName}" deleted successfully.`);
        } else {
            throw new Error(`Failed to delete folder "${folderName}"`);
        }
    } catch (error) {
       // console.error(error.message);
        throw error; // Rethrow the error to handle it in the endpoint
    }
}

// Function to create a new folder
async function createFolder(token, parentFolderId, folderName) {
    try {
        const response = await axios.post(
            `https://graph.microsoft.com/v1.0/me/mailFolders`,
            {
                displayName: folderName,
                parentFolderId: parentFolderId // Use the fixed parentFolderId
            },
            {
                headers: {
                    Authorization: `Bearer ${token}`,
                    'Content-Type': 'application/json'
                }
            }
        );
        console.log('Folder created:', response.data);
    } catch (error) {
        console.error('Error creating folder:', error);
        throw error; // Rethrow the error to handle it in the endpoint
    }
}

// Function to rename a folder
async function renameFolder(token, folderId, newFolderName) {
    try {
        const response = await axios.patch(
            `https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}`,
            {
                displayName: newFolderName
            },
            {
                headers: {
                    Authorization: `Bearer ${token}`,
                    'Content-Type': 'application/json'
                }
            }
        );
        console.log('Folder renamed:', response.data);
    } catch (error) {
        console.error('Error renaming folder:', error);
        throw error; // Rethrow the error to handle it in the endpoint
    }
}

// Endpoint to generate email using AI
app.post('/api/compose', async (req, res) => {
    const { accessToken, body } = req.body; // Get access token, subject, and body from request body
   console.log(body)
    if (!accessToken  || !body) {
        return res.status(400).json({ error: 'Access token, subject, and body are required' });
    }

    try {
        // Construct the prompt for AI
        const prompt = `You're a Smart Email Composer API, Compose a professional email using prompt: ${body}`;

        // Generate the email content
        const response = await model.generateContent(prompt); // Use generateContent instead of invoke
        // Extract the generated email content
        const result=response.response.text();
        console.log('Response:',result)

        res.status(200).json({ email:result });
        
    } catch (error) {
        console.error('Error generating email:', error);
        return res.status(500).json({ error: 'Failed to generate email' });
    }
});


//email sentiment categorization endpoint
app.post('/api/sentiment', async (req, res) => {
    try {
        const { emailContent } = req.body;
        let {emailSubject} = req.body;
        console.log('Sentiment Section');
        console.log(emailContent);

        // Validate input
        if (!emailContent || typeof emailContent !== 'string') {
            return res.status(400).json({ error: 'Invalid email content provided' });
        }

        const emailBody = emailContent.trim(); // Replace with your `maskPII` logic if needed
        emailSubject = emailSubject.trim();
        //console.log('Masked Text:', maskedText);

        // Create the prompt template for sentiment analysis
        const promptTemplate = `
            Analyze the following email and respond with a JSON object in the exact format provided below, also make sure there are no leading/trailing characters (e.g., whitespace, quotes, or unexpected symbols in between or end or start of json. The output must not include any additional text, code blocks, or formatting. Ensure the JSON is valid and does not include comments or extraneous characters.

            Format:
            {
                "Priority Level":High/Medium/Low,
                "Urgency":High/Medium/Low,
                "Sentiment":Positive/Neutral/Negative,
                "Category":Complaint/Query/Feedback/Product Enquiry/etc,
                "Impact":High/Medium/Low,
                "Product":[]
            }

            Here is the email Subject:
            ${emailSubject}

            Here is the email Body:
            ${emailBody}

            If the category is "Product Enquiry", based on the email content, fill the "Product" array with the closest matches to at least one of the products listed below. If no match is found, leave the "Product" array empty.

            Products:
            {
                "Managed IT Support Services",
                "Cloud Hosting Solutions",
                "Custom Software Development",
                "IT Consulting",
                "Cybersecurity Solutions",
                "IT Infrastructure Setup",
                "Backup and Disaster Recovery",
                "IT Audits and Compliance",
                "Network Design and Implementation",
                "Hardware Procurement and Installation"
            }
    
        `;

        // Create a chat message format
        const messages = [{ role: 'user', content: promptTemplate }];

        const model = getLlmInstance();

        // Invoke the model with the chat messages
        const response = await model.invoke(messages);
        console.log('the sentiment analysis content is ***********************************');
        console.log(response.content);  
        // if(response.content.Category == 'Product Enquiry')
        // {
        //     response.content['Product Enquired'] = 
        // }
        // Send the response directly
        res.status(200).json(response);
    } catch (error) {
        console.error('Error during sentiment analysis:', error);
        res.status(500).json({ error: 'Failed to analyze email sentiment' });
    }
});



// Add new endpoint to create subscription for email notifications
app.post('/api/createEmailSubscription', async (req, res) => {
    const { accessToken } = req.body;

    if (!accessToken) {
        return res.status(401).json({ message: 'Access token is required' });
    }

    try {
        const subscription = {
            changeType: 'created',
            notificationUrl: `${process.env.BACKEND_URL}/api/notifications`, // Your backend webhook URL
            resource: 'me/messages',
            expirationDateTime: new Date(Date.now() + 60 * 60 * 24 * 1000).toISOString(), // 24 hours
            clientState: 'secretClientValue' // Add this to your .env file
        };

        const response = await axios.post(
            'https://graph.microsoft.com/v1.0/subscriptions',
            subscription,
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            }
        );

       // console.log('Subscription created:', response.data);
        res.status(200).json({ subscription: response.data });
    } catch (error) {
        //console.error('Error creating subscription:', error);
        res.status(500).json({ message: 'Error creating email subscription' });
    }
});


// Endpoint to renew subscription
app.post('/api/renewSubscription', async (req, res) => {
    const { accessToken, subscriptionId } = req.body;

    if (!accessToken || !subscriptionId) {
        return res.status(400).json({ message: 'Access token and subscription ID are required' });
    }

    try {
        const newExpirationDateTime = new Date(Date.now() + 60 * 60 * 24 * 1000).toISOString(); // 24 hours

        const response = await axios.patch(
            `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`,
            {
                expirationDateTime: newExpirationDateTime
            },
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            }
        );

        //console.log('Subscription renewed:', response.data);
        res.status(200).json({ subscription: response.data });
    } catch (error) {
        //console.error('Error renewing subscription:', error);
        res.status(500).json({ message: 'Error renewing subscription' });
    }
});

// Endpoint to delete subscription
app.delete('/api/subscription/:subscriptionId', async (req, res) => {
    const { accessToken } = req.body;
    const { subscriptionId } = req.params;

    if (!accessToken || !subscriptionId) {
        return res.status(400).json({ message: 'Access token and subscription ID are required' });
    }

    try {
        await axios.delete(
            `https://graph.microsoft.com/v1.0/subscriptions/${subscriptionId}`,
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`
                }
            }
        );

        res.status(200).json({ message: 'Subscription deleted successfully' });
    } catch (error) {
        //console.error('Error deleting subscription:', error);
        res.status(500).json({ message: 'Error deleting subscription' });
    }
});

// Endpoint to update the read status of a specific email
app.patch('/api/markAsRead', async (req, res) => {
    const { accessToken,messageId } = req.body; // Get the access token from the request body
    console.log('Mark As Read API')
    if (!accessToken) {
        return res.status(401).json({ message: 'Access token is required' });
    }

    if (!messageId) {
        return res.status(400).json({ message: 'Email ID is required' });
    }

    try {
        // Update the read status of the email
        await axios.patch(
            `https://graph.microsoft.com/v1.0/me/messages/${messageId}`,
            { isRead: true },
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            }
        );

        console.log(`Email with ID ${messageId} marked as read.`);
        return res.status(200).json({ message: `Email with ID ${messageId} marked as read.` });
    } catch (error) {
        console.error('Error updating email status:', error.response ? error.response.data : error.message);
        return res.status(500).json({ message: 'Error updating email status.' });
    }
});

// salesforce product 

app.post('/api/fetchProductDetails' , async(req , res) =>{
    try{
        const {products} =  req.body;
        const username = 'joiningapp7@resourceful-fox-7w3tc4.com'; // Salesforce username
        const password = 'Hello@123123CwozoRmQ28h4Fkx24mCyNU8q'; // Salesforce password
        const connection = new jsforce.Connection({
            loginUrl: 'https://login.salesforce.com'
        });
        try {
            await connection.login(username, password); // Log in to Salesforce
            console.log('Logged in successfully!');
            const endpoint = `/ProductSummary?products=${products}`;
            const response = await connection.apex.get(endpoint);
            console.log('Response:', response);
            res.status(200).json({ ...response }); 
            // Send details back to client
        } catch (error) {
            console.error('Error during Salesforce operation:', error);
            return res.status(500).json({ message: 'Error fetching product details from Salesforce' });
        }

    }
    catch(exception)
    {

    }
});



// Start the server
const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});






async function getEmailInsights(emailContent) {
  // Validate input
  if (!emailContent || typeof emailContent !== 'string') {
    throw new Error('Invalid email content provided');
  }

  const emailInsightPrompt = new PromptTemplate({
    template: `Analyze the following email and provide a detailed priority assessment based on the following factors:
              Urgency: Does the email mention deadlines, time-sensitive issues, or immediate consequences?
              Sentiment: Assess the tone as Positive, Neutral, or Negative and indicate its impact on urgency.
              Category: Classify the email into one of the following groups: Complaint, Query, Feedback, Request, Follow-Up, Escalation, or Other.
              Impact: Determine the potential effect of the email on the organization (e.g., financial, reputational, operational).
              Provide your response in the following structured format:

              Urgency: [High/Medium/Low]
              Sentiment: [Positive/Neutral/Negative]
              Category: [Complaint/Query/Feedback/etc.]
              Impact: [High/Medium/Low]
              Priority Level: [High/Medium/Low]
              Here is the email: 
              ${emailContent}`,
    inputVariables: ["emailContent"]
  });

  try {
    const formattedPrompt = await emailInsightPrompt.format({
      emailContent: emailContent
    });
    console.log(formattedPrompt);
    const response = await model.invoke(formattedPrompt);
    
    return response.content;
  } catch (error) {
    //console.error('failed:', error);
    throw new Error('Failed to get insights of an email');
  }
}

