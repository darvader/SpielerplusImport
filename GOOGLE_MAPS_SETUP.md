# Google API Setup Guide

## Overview
The SpielerplusImport script can use **one Google API key** for two services:
1. **Distance Matrix API** - Real-time travel calculations
2. **Cloud Translation API** - Advanced German text encoding fixes

## Setup Steps

### 1. Create Google Cloud Project
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing one
3. **Set up billing** (required for both APIs)

### 2. Enable Required APIs
1. In Google Cloud Console, go to **APIs & Services** > **Library**
2. Search and enable both:
   - **"Distance Matrix API"** (for travel times)
   - **"Cloud Translation API"** (for text encoding)

### 3. Create Single API Key
1. Go to **APIs & Services** > **Credentials**
2. Click **Create Credentials** > **API Key**
3. Copy the generated API key
4. (Optional) Set restrictions:
   - **API restrictions**: Select only the two APIs above

### 4. Configure the Script
1. Edit your `.env` file:
   ```properties
   GOOGLE_API_KEY=AIzaSyYourActualApiKeyHere
   ```

## What Each Service Does

### Distance Matrix API
- **Purpose**: Calculate real-time travel times between cities
- **Fallback**: Static travel time estimates if API fails
- **Cost**: $0.005 per request (2,500 free per day)

### Cloud Translation API  
- **Purpose**: Fix German encoding issues like "Oberwei�bach" → "Oberweißbach"
- **Fallback**: Local pattern matching if API fails
- **Cost**: $20 per 1M characters (500,000 free per month)

## Cost Information

**For typical volleyball schedule processing**:
- **Distance Matrix**: ~10-20 requests per import = Well within free tier
- **Translation**: ~1,000 characters per import = Well within free tier

**Monthly costs for active teams**: $0 (free tier sufficient)

## When APIs Are Used

### Travel Time Calculation
```
✅ Google API: Real-time traffic and route conditions
❌ Fallback: Static estimates (Jena→Erfurt = 60 min)
```

### Text Encoding
```
✅ Google API: "Oberwei�bach" detected and fixed automatically  
❌ Fallback: Pattern matching for known issues
```

## Testing Your Setup

Run the script and look for these messages:

### ✅ **Success**:
```
Configuration:
  Google API: Configured
  
Calculating travel time from 07747 Jena to 98617 Meiningen using Google Maps...
Google Maps result: 89.2 km, 87 minutes

Using Google Cloud Translation API to fix encoding for: Oberwei�bach
Google API cleaned text: Oberweißbach
```

### ❌ **API Issues**:
```
WARNING: Google Maps API REQUEST_DENIED
WARNING: Google Cloud Translation API failed
```

## Troubleshooting

### REQUEST_DENIED Error
- ✅ Both APIs are enabled in Google Cloud Console
- ✅ Billing is set up and active
- ✅ API key has no conflicting restrictions
- ✅ API key is not expired

### OVER_QUERY_LIMIT Error
- Check quotas in Google Cloud Console
- Increase limits if needed (usually not required)

### Cost Monitoring
1. Go to **Billing** in Google Cloud Console
2. Set up billing alerts
3. Monitor API usage in **APIs & Services** > **Quotas**

## Security Best Practices

- ✅ API key stored in `.env` file (excluded from git)
- ✅ Consider IP restrictions for production use
- ✅ Monitor usage and set billing alerts
- ✅ Use least-privilege API restrictions

## Alternative: Skip Google APIs

The script works perfectly without Google APIs:
- Uses static travel time estimates
- Uses local German character fixes
- No external dependencies or costs
- Processes data offline
