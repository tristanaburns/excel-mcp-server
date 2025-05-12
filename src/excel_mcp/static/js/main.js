// Main JavaScript file for Excel MCP Server

document.addEventListener('DOMContentLoaded', function() {
    // Auto-dismiss alerts after 5 seconds
    const alerts = document.querySelectorAll('.alert');
    alerts.forEach(alert => {
        setTimeout(() => {
            const bsAlert = new bootstrap.Alert(alert);
            bsAlert.close();
        }, 5000);
    });

    // Active navigation highlighting
    const currentPath = window.location.pathname;
    const navLinks = document.querySelectorAll('.navbar-nav .nav-link');
    
    navLinks.forEach(link => {
        if (link.getAttribute('href') === currentPath) {
            link.classList.add('active');
        }
    });    // Responsive files container height adjustment
    function adjustFilesContainerHeight() {
        const filesContainer = document.querySelector('.files-container');
        const cardBody = document.querySelector('.card-body');
        const card = document.querySelector('.card');
        const mainContent = document.querySelector('main');
        
        if (filesContainer && cardBody && card && mainContent) {
            const windowHeight = window.innerHeight;
            const navHeight = document.querySelector('header')?.offsetHeight || 70;
            const footerHeight = document.querySelector('footer')?.offsetHeight || 40;
            const pageHeaderHeight = document.querySelector('.d-flex.justify-content-between')?.offsetHeight || 0;
            const alert = document.querySelector('.alert');
            const alertHeight = alert ? alert.offsetHeight + 16 : 0; // 16px for margin
            
            // Calculate available space
            const availableHeight = windowHeight - navHeight - footerHeight - 40; // 40px extra padding
            
            // Set main content min-height
            mainContent.style.minHeight = `${availableHeight}px`;
            
            // Calculate card height allowing for page elements
            const cardHeight = availableHeight - pageHeaderHeight - alertHeight - 40; // 40px for padding
            
            // Set container height (allowing space for dropdowns)
            const containerHeight = cardHeight - 120; // Space for dropdowns
            
            // Apply heights
            card.style.minHeight = `${cardHeight}px`;
            filesContainer.style.maxHeight = `${containerHeight}px`;
            
            // Add margin at bottom of table based on number of rows
            const tableFiles = document.querySelector('.table-files');
            if (tableFiles) {
                const numRows = document.querySelectorAll('.file-row').length;
                const extraMargin = Math.max(60, numRows * 15); // At least 60px or 15px per row
                tableFiles.style.marginBottom = `${extraMargin}px`;
            }
        }
    }
    // Handle dropdown positioning to prevent overflow
    function setupDropdownPositioning() {
        const dropdownToggles = document.querySelectorAll('.dropdown-toggle');
        
        // First, ensure all dropdowns have the right classes
        document.querySelectorAll('.dropdown-menu').forEach(menu => {
            menu.classList.add('dropdown-menu-end');
        });
        
        // Then set up click handlers for each dropdown toggle
        dropdownToggles.forEach(toggle => {
            toggle.addEventListener('click', function(e) {
                // Prevent immediate closing of dropdown
                e.stopPropagation();
                
                // Small delay to allow dropdown to render
                setTimeout(() => {
                    const menu = this.nextElementSibling;
                    if (!menu || !menu.classList.contains('dropdown-menu')) return;
                    
                    // Get viewport dimensions
                    const viewportWidth = window.innerWidth;
                    const viewportHeight = window.innerHeight;
                    
                    // Get dropdown position
                    const rect = menu.getBoundingClientRect();
                    const toggleRect = toggle.getBoundingClientRect();
                    
                    // Calculate the space available below the toggle
                    const spaceBelow = viewportHeight - toggleRect.bottom;
                    const spaceAbove = toggleRect.top;
                    
                    // Reset dropdown positioning
                    menu.style.top = '';
                    menu.style.bottom = '';
                    menu.style.left = '';
                    menu.style.right = '';
                    menu.style.transform = '';
                    
                    // If dropdown extends beyond bottom of viewport and there's more space above
                    if (rect.height > spaceBelow && spaceAbove > spaceBelow) {
                        menu.style.top = 'auto';
                        menu.style.bottom = '100%';
                    }
                    
                    // Ensure dropdown is visible horizontally by keeping it within viewport
                    if (rect.right > viewportWidth) {
                        const overflow = rect.right - viewportWidth;
                        menu.style.right = '0';
                        menu.style.left = 'auto';
                    }
                    
                    // Add extra space at the bottom of the container if needed
                    const filesContainer = document.querySelector('.files-container');
                    if (filesContainer) {
                        const numRows = document.querySelectorAll('.file-row').length;
                        const extraSpace = Math.max(80, numRows * 20); // At least 80px or 20px per row
                        filesContainer.style.paddingBottom = `${extraSpace}px`;
                    }
                }, 20);
                        });
        });
    }
    
    // Function to handle table scrolling
    function setupTableScrolling() {
        const filesContainer = document.querySelector('.files-container');
        if (!filesContainer) return;
        
        // Ensure the table has enough room to display dropdowns when scrolled
        filesContainer.addEventListener('scroll', () => {
            // Add more bottom padding when scrolled near the bottom
            const scrollPos = filesContainer.scrollTop;
            const scrollHeight = filesContainer.scrollHeight;
            const clientHeight = filesContainer.clientHeight;
            
            // If scrolled more than 70% down, ensure extra padding
            if (scrollPos + clientHeight > scrollHeight * 0.7) {
                const tableFiles = document.querySelector('.table-files');
                if (tableFiles) {
                    const currentMargin = parseInt(tableFiles.style.marginBottom || '60');
                    if (currentMargin < 120) {
                        tableFiles.style.marginBottom = '120px';
                    }
                }
            }
        });
    }
    
    // Function to handle modal events
    function setupModalHandlers() {
        // Reset container heights when modals open/close
        const modals = document.querySelectorAll('.modal');
        modals.forEach(modal => {
            modal.addEventListener('hidden.bs.modal', () => {
                setTimeout(adjustFilesContainerHeight, 100);
            });
            
            modal.addEventListener('shown.bs.modal', () => {
                // Close any open dropdowns when a modal opens
                document.querySelectorAll('.dropdown-menu.show').forEach(menu => {
                    const dropdownToggle = menu.previousElementSibling;
                    if (dropdownToggle) {
                        const bsDropdown = bootstrap.Dropdown.getInstance(dropdownToggle);
                        if (bsDropdown) bsDropdown.hide();
                    }
                });
            });
        });
    }
    
    // Initial adjustments
    adjustFilesContainerHeight();
    setupDropdownPositioning();
    setupTableScrolling();
    setupModalHandlers();
    
    // Adjust on window resize
    window.addEventListener('resize', () => {
        adjustFilesContainerHeight();
    });
    
    // Add a small delay to make sure everything is rendered properly
    setTimeout(() => {
        adjustFilesContainerHeight();
        setupDropdownPositioning();
    }, 200);
});
