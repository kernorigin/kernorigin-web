// ============================================
// KERNORIGIN MOBILE MENU HANDLER
// Production v2.0 - Ready to Deploy
// ============================================

(function() {
    'use strict';
    
    // Create mobile menu toggle button
    function initMobileMenu() {
        const nav = document.querySelector('nav');
        if (!nav) return;
        
        // Create hamburger button
        const hamburger = document.createElement('button');
        hamburger.className = 'mobile-menu-toggle';
        hamburger.setAttribute('aria-label', 'Toggle navigation');
        hamburger.innerHTML = `
            <span class="hamburger-line"></span>
            <span class="hamburger-line"></span>
            <span class="hamburger-line"></span>
        `;
        
        // Insert hamburger before nav
        nav.parentElement.insertBefore(hamburger, nav);
        
        // Toggle menu on click
        hamburger.addEventListener('click', function() {
            nav.classList.toggle('mobile-menu-open');
            hamburger.classList.toggle('active');
            document.body.classList.toggle('menu-open');
        });
        
        // Close menu when clicking outside
        document.addEventListener('click', function(e) {
            if (!nav.contains(e.target) && !hamburger.contains(e.target)) {
                nav.classList.remove('mobile-menu-open');
                hamburger.classList.remove('active');
                document.body.classList.remove('menu-open');
            }
        });
        
        // Close menu when clicking on a link
        const navLinks = nav.querySelectorAll('a');
        navLinks.forEach(link => {
            link.addEventListener('click', function() {
                nav.classList.remove('mobile-menu-open');
                hamburger.classList.remove('active');
                document.body.classList.remove('menu-open');
            });
        });
    }
    
    // Initialize when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initMobileMenu);
    } else {
        initMobileMenu();
    }
})();
