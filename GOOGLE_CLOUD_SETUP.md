# Google Cloud Translation API Setup (Optional)

## Overview
The script can optionally use Google Cloud Translation API for more sophisticated German text encoding fixes. This is particularly useful if your CSV files have complex encoding issues that the built-in fixes don't handle.

**Note**: For most volleyball schedules, the built-in encoding fixes are sufficient and free. Use Google Cloud API only if you have persistent encoding issues.

## Setup Steps

### 1. Create Google Cloud Project
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing one
3. Enable billing (required for API access)

### 2. Enable Cloud Translation API
1. In Google Cloud Console, go to **APIs & Services** > **Library**
2. Search for "Cloud Translation API"
3. Click on it and press **Enable**

### 3. Create API Key
1. Go to **APIs & Services** > **Credentials**
2. Click **Create Credentials** > **API Key**
3. Copy the generated API key
4. (Optional) Set restrictions:
   - **Application restrictions**: None or IP addresses
   - **API restrictions**: Restrict to Cloud Translation API

### 4. Configure the Script
1. Edit your `.env` file:
   ```properties
   GOOGLE_CLOUD_API_KEY=AIzaSyYourActualCloudApiKeyHere
   ```

## How It Works

### Intelligent Fallback System
1. **First**: Tries Google Cloud Translation API if configured and text contains '�' characters
2. **Fallback**: Uses local pattern matching rules if API fails or isn't configured

### Example Usage
```
Input:  "Oberwei�bach Th�ringenliga"
Google: "Oberweißbach Thüringenliga"
Local:  "Oberweißbach Thüringenliga"
```

## Cost Information

Google Cloud Translation API pricing:
- **Free tier**: First 500,000 characters per month
- **Paid tier**: $20 per 1M characters

For volleyball schedules:
- Typical usage: ~1,000 characters per schedule
- Well within free tier for most teams

## When to Use Google Cloud API

### ✅ **Use Google Cloud API if**:
- You have persistent encoding issues with team names
- CSV files come from different sources with varying encodings
- You need to process multiple languages
- Built-in fixes don't cover your specific cases

### ❌ **Skip Google Cloud API if**:
- Built-in fixes work fine for your data
- You want to minimize external dependencies
- You prefer offline processing
- Cost is a concern

## Monitoring Usage

1. Check usage in Google Cloud Console:
   - Go to **APIs & Services** > **Quotas**
   - Monitor "Cloud Translation API" usage

2. Script output shows when API is used:
   ```
   Using Google Cloud Translation API to fix encoding for: Oberwei�bach
   Google API cleaned text: Oberweißbach
   ```

## Troubleshooting

### API Key Issues
- Ensure Cloud Translation API is enabled
- Check API key restrictions
- Verify billing is set up

### Rate Limiting
- Google Cloud has generous limits
- Script processes text sequentially to avoid issues

### Fallback Behavior
- If API fails, script automatically uses local fixes
- No data loss if API is unavailable

## Security Notes

- API key is stored in `.env` file (excluded from git)
- Consider using service accounts for production
- Monitor API usage and costs
- Set up billing alerts in Google Cloud Console
