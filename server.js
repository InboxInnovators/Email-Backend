import express from 'express';
import cors from 'cors';
import axios from 'axios';
import 'dotenv/config';
import path from 'path';
import fs from 'fs';

import { fileURLToPath } from 'url';
import { dirname } from 'path';
import { GoogleGenerativeAI } from '@google/generative-ai';
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const app = express();

// Initialize Google Generative AI
const genAI = new GoogleGenerativeAI('AIzaSyAF0TUxoaLIV9hQbCKj6jlUFyXA64KecG8');
const model = genAI.getGenerativeModel({ model: 'gemini-1.5-flash' });

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
        const response = await axios.get("https://graph.microsoft.com/v1.0/me/messages", {
            headers: {
                Authorization: `Bearer ${accessToken}`,
            },
        });
        
        const emails = response.data.value; // Extract the emails
        res.status(200).json({ emails });
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
        console.log('Fetched folders:', response.data);

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
        console.error('Error fetching folders:', error);
        return res.status(500).json({ error: 'Failed to fetch folders' });
    }
});

//Endpoint to summarize the email using AI
app.post('/api/summarize', async (req, res) => {
    try {
        const { body } = req.body;

        // Validate email content
        if (!body || typeof body !== 'string') {
            return res.status(400).json({ error: 'Invalid email content provided' });
        }

        // Construct the prompt
        const prompt = `Summarize the following email in the most concise way possible, focusing on the key points such as the purpose of the email, any actions required, deadlines, or important details. Adjust the length of the summary based on the email's content; if the email is short, provide a proportionately brief summary. `;
        const userInput = body;

        // Generate summary using Google Generative AI
        const result = await model.generateContent(prompt + userInput);

        // Send the summary back to the client
        const summary = result.response.text(); // Ensure this is correct
        return res.status(200).json({ summary });
    } catch (error) {
        console.error('Error summarizing email:', error);
        return res.status(500).json({ error: 'Failed to summarize email' });
    }
});


// New endpoint to send and save the email
app.post('/api/sendEmail', async (req, res) => {
    const { accessToken, subject, body, recipients } = req.body; // Get access token and email details from request body

    if (!accessToken) {
        return res.status(401).json({ message: 'Access token is required' });
    }

    if (!subject || !body || !recipients || !Array.isArray(recipients)) {
        return res.status(400).json({ message: 'Subject, body, and recipients are required' });
    }

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
        }))
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
        const response = await model.generateContent(`Translate the following text from ${sourceLanguage} to ${targetLanguage}: ${text}`);
        const result=response.response.text();
        console.log('Response:',result)

        res.json({ result });
    } catch (error) {
        console.error('Error during translation:', error);
        res.status(500).json({ error: 'Translation failed' });
    }
});


// Start the server
const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
