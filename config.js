// ============================================
// KERNORIGIN FRONTEND CONFIG
// Production v2.0 - Ready to Deploy
// ============================================
// SETUP INSTRUCTIONS:
// 1. Replace YOUR_SCRIPT_ID with your actual Google Apps Script deployment ID
// 2. Replace YOUR_PUBLIC_API_KEY with your actual key from Script Properties
// 3. Upload this file to your web hosting alongside index.html
// ============================================

const API_CONFIG = {
    // Your Google Apps Script Web App URL
    // Must end in /exec
    URL: 'https://script.google.com/macros/s/AKfycbypfCu9RFTq7o9FZ_VvRx_QfE9UnFyBwyfUzrdPlSHe9dzwcxy7wCGf2htWzmXmnNoO/exec',
    
    // Your PUBLIC_API_KEY from Script Properties
    PUBLIC_KEY: 'Kernorigin-Public-2028-ANJANA'
};

// ============================================
// CONFIGURATION STATUS CHECK
// ============================================
// This config.js is ready to use!
// 
// ✓ URL is configured (your deployment URL is set)
// ⚠  PUBLIC_KEY needs to be verified
//
// To verify PUBLIC_KEY:
// 1. Open your AppScript project
// 2. Go to Project Settings → Script Properties
// 3. Find PUBLIC_API_KEY
// 4. Make sure it matches the value above
//
// If you need to update PUBLIC_KEY, just replace
// 'Kernorigin-Public-2028-ANJANA' with your actual key
// ============================================

// Validation on load
(function validateConfig() {
    if (API_CONFIG.URL.includes('YOUR_SCRIPT_ID')) {
        console.warn('⚠️ API_CONFIG.URL not configured - update config.js with your deployment URL');
    } else {
        console.log('✓ API URL configured');
    }
    
    if (API_CONFIG.PUBLIC_KEY === 'YOUR_PUBLIC_API_KEY' || API_CONFIG.PUBLIC_KEY.includes('XYZ123')) {
        console.warn('⚠️ API_CONFIG.PUBLIC_KEY may need updating - verify it matches Script Properties');
    } else {
        console.log('✓ PUBLIC_KEY set (verify it matches Script Properties)');
    }
})();