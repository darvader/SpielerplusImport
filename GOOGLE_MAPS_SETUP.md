# Google Maps API Setup Guide

## Overview
The SpielerplusImport script can use Google Maps Distance Matrix API to calculate real-time travel times between venues. This provides more accurate meeting time calculations than the built-in static estimates.

## Setup Steps

### 1. Create Google Cloud Project
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing one
3. Note your project ID

### 2. Enable Distance Matrix API
1. In Google Cloud Console, go to **APIs & Services** > **Library**
2. Search for "Distance Matrix API"
3. Click on it and press **Enable**

### 3. Set up Billing
1. Go to **Billing** in Google Cloud Console
2. Link a payment method to your project
3. **Note**: Google Maps API requires billing even for free tier usage

### 4. Create API Key
1. Go to **APIs & Services** > **Credentials**
2. Click **Create Credentials** > **API Key**
3. Copy the generated API key
4. (Optional) Click on the key to set restrictions:
   - **Application restrictions**: None or HTTP referrers
   - **API restrictions**: Restrict to Distance Matrix API

### 5. Configure the Script
1. Copy `.env.example` to `.env`:
   ```powershell
   Copy-Item .env.example .env
   ```
2. Edit `.env` file and replace `your_google_maps_api_key_here` with your actual API key:
   ```
   GOOGLE_MAPS_API_KEY=AIzaSyYourActualApiKeyHere
   ```

### 6. Test the Setup
Run the script and check for successful Google Maps calls:
```powershell
.\Transform-Spielplan-Simple.ps1
```

Look for messages like:
- ✅ `Google Maps result: 45.2 km, 67 minutes`
- ❌ `WARNING: Google Maps API REQUEST_DENIED`

## Common Issues

### REQUEST_DENIED Error
This usually means:
- Distance Matrix API is not enabled
- Billing is not set up
- API key has restrictions blocking the request
- API key is invalid or expired

### OVER_QUERY_LIMIT Error
- You've exceeded the free tier limits
- Need to increase quotas or billing limits

### ZERO_RESULTS Error
- The addresses couldn't be found
- Try different address formats

## Cost Information

Google Maps Distance Matrix API pricing (as of 2025):
- **Free tier**: 2,500 requests per day
- **Paid tier**: $0.005 per request after free tier

For volleyball schedule processing:
- Typical usage: 10-20 requests per schedule import
- Well within free tier limits for most teams

## Fallback Behavior

If Google Maps API fails or is not configured, the script automatically falls back to static travel time estimates:
- Jena ↔ Erfurt: 60 minutes
- Jena ↔ Weimar: 45 minutes  
- Jena ↔ Gera: 45 minutes
- Jena ↔ Meiningen: 90 minutes
- And more...

These estimates are based on typical driving times between Thuringian cities.

## Security Notes

- The `.env` file is automatically added to `.gitignore`
- Never commit API keys to version control
- Consider using API key restrictions for additional security
- Monitor your Google Cloud billing and usage
