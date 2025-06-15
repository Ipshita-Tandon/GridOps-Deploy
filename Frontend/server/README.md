# WOPI Server Setup

This server provides WOPI (Web Application Open Platform Interface) functionality for viewing Excel files in the browser using Microsoft Office Online.

## Prerequisites

1. Node.js (v14 or higher)
2. Python 3.7 or higher
3. ngrok (for development)

## Installation

1. Install Node.js dependencies:

```bash
npm install
```

2. Install Python dependencies:

```bash
pip install -r requirements.txt
```

3. Create a `public` directory to store Excel files:

```bash
mkdir public
```

4. Copy your Excel files to the `public` directory.

## Running the Server

1. Start the WOPI server:

```bash
npm start
```

2. Start ngrok (in a separate terminal):

```bash
ngrok http 3000
```

The server will be available at `http://localhost:3000` and through the ngrok URL.

## Important Notes

1. The server requires ngrok for Microsoft Office Online to access it. Make sure ngrok is running before using the Excel viewer.
2. Excel files should be placed in the `public` directory.
3. The server uses a hardcoded access token ("12345") for development purposes. In production, implement proper authentication.
4. Cell highlighting requires the Python script to be running and accessible.
