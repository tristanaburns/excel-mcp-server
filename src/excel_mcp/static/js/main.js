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
    });
    
    // Responsive files container height adjustment
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
                const extraMargin = Math.max(80, numRows * 20); // At least 80px or 20px per row
                tableFiles.style.marginBottom = `${extraMargin}px`;
            }
        }
    }
    
    // Enhanced dropdown positioning to prevent overflow
    function setupDropdownPositioning() {
        // Clean up any existing event listeners to prevent duplicates
        document.querySelectorAll('.dropdown-toggle').forEach(toggle => {
            const newToggle = toggle.cloneNode(true);
            toggle.parentNode.replaceChild(newToggle, toggle);
        });
        
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
                    
                    // Apply positioning based on available space
                    if (rect.height > spaceBelow && spaceAbove > spaceBelow) {
                        // If there's more space above, show dropdown above the button
                        menu.style.top = 'auto';
                        menu.style.bottom = '100%';
                        menu.style.marginBottom = '2px';
                        menu.style.marginTop = '0';
                    } else {
                        // Otherwise, show below with enough margin
                        menu.style.top = '100%';
                        menu.style.bottom = 'auto';
                        menu.style.marginTop = '2px';
                        menu.style.marginBottom = '0';
                    }
                    
                    // Always ensure horizontal alignment
                    menu.style.right = '0';
                    menu.style.left = 'auto';
                    
                    // Ensure the menu is within the viewport width
                    if (rect.right > viewportWidth) {
                        const overflow = rect.right - viewportWidth;
                        menu.style.right = `${overflow + 10}px`; // 10px buffer
                    }
                    
                    // Add extra space at the bottom of the container
                    const filesContainer = document.querySelector('.files-container');
                    const tableFiles = document.querySelector('.table-files');
                    
                    if (filesContainer && tableFiles) {
                        const numRows = document.querySelectorAll('.file-row').length;
                        const extraSpace = Math.max(100, numRows * 25); // At least 100px or 25px per row
                        
                        if (rect.bottom > viewportHeight - 50) {
                            // If dropdown extends beyond viewport, add even more space
                            tableFiles.style.marginBottom = `${extraSpace + 50}px`;
                        } else {
                            tableFiles.style.marginBottom = `${extraSpace}px`;
                        }
                    }
                }, 20);
            });
        });
        
        // Close other dropdowns when one is opened
        document.addEventListener('click', function(e) {
            const openDropdowns = document.querySelectorAll('.dropdown-menu.show');
            
            openDropdowns.forEach(dropdown => {
                const toggle = dropdown.previousElementSibling;
                if (!toggle) return;
                
                if (!toggle.contains(e.target) && !dropdown.contains(e.target)) {
                    const bsDropdown = bootstrap.Dropdown.getInstance(toggle);
                    if (bsDropdown) bsDropdown.hide();
                }
            });
        });
    }
    
    // Function to handle table scrolling for dropdowns
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
                    const currentMargin = parseInt(tableFiles.style.marginBottom || '80');
                    if (currentMargin < 150) {
                        tableFiles.style.marginBottom = '150px';
                    }
                }
            }
        });
    }
    
    // Function to handle multiple file uploads
    function setupFileUpload() {
        const uploadForm = document.getElementById('uploadForm');
        const uploadBtn = document.getElementById('uploadBtn');
        const fileInput = document.getElementById('excelFile');
        const uploadProgress = document.getElementById('uploadProgress');
        const progressBar = uploadProgress.querySelector('.progress-bar');
        const uploadStatus = document.getElementById('uploadStatus');
        const uploadModal = document.getElementById('uploadModal');
        
        if (uploadBtn && uploadForm && fileInput) {
            uploadBtn.addEventListener('click', function() {
                // Validate file selection
                if (!fileInput.files || fileInput.files.length === 0) {
                    alert('Please select at least one Excel file to upload.');
                    return;
                }
                
                // Check file types
                for (const file of fileInput.files) {
                    const extension = file.name.split('.').pop().toLowerCase();
                    if (extension !== 'xlsx' && extension !== 'xls') {
                        alert(`File "${file.name}" is not a valid Excel file. Only .xlsx and .xls files are supported.`);
                        return;
                    }
                }
                
                // Show progress bar
                uploadProgress.classList.remove('d-none');
                progressBar.style.width = '0%';
                uploadStatus.innerHTML = '';
                
                // Create FormData for file upload
                const formData = new FormData();
                for (const file of fileInput.files) {
                    formData.append('file', file);
                }
                
                // Make AJAX request to upload files
                const xhr = new XMLHttpRequest();
                xhr.open('POST', '/upload', true);
                
                // Track upload progress
                xhr.upload.addEventListener('progress', function(e) {
                    if (e.lengthComputable) {
                        const percent = Math.round((e.loaded / e.total) * 100);
                        progressBar.style.width = percent + '%';
                        progressBar.textContent = percent + '%';
                    }
                });
                
                // Handle response
                xhr.onload = function() {
                    if (xhr.status >= 200 && xhr.status < 400) {
                        // Success - refresh page to show new files
                        window.location.reload();
                    } else {
                        // Error
                        uploadStatus.innerHTML = `
                            <div class="alert alert-danger">
                                Upload failed. Please try again. (${xhr.status}: ${xhr.statusText})
                            </div>
                        `;
                        uploadProgress.classList.add('d-none');
                    }
                };
                
                xhr.onerror = function() {
                    uploadStatus.innerHTML = `
                        <div class="alert alert-danger">
                            Network error occurred. Please check your connection and try again.
                        </div>
                    `;
                    uploadProgress.classList.add('d-none');
                };
                
                // Send the request
                xhr.send(formData);
            });
            
            // Reset form when modal is closed
            if (uploadModal) {
                uploadModal.addEventListener('hidden.bs.modal', function() {
                    uploadForm.reset();
                    uploadProgress.classList.add('d-none');
                    uploadStatus.innerHTML = '';
                });
            }
        }
    }
    
    // Function to handle file info modal
    function setupInfoModal() {
        const infoModal = document.getElementById('infoModal');
        
        if (infoModal) {
            infoModal.addEventListener('show.bs.modal', function(event) {
                const button = event.relatedTarget;
                const filename = button.getAttribute('data-filename');
                
                const infoLoading = document.getElementById('infoLoading');
                const infoContent = document.getElementById('infoContent');
                const infoError = document.getElementById('infoError');
                
                // Show loading, hide content and error
                infoLoading.classList.remove('d-none');
                infoContent.classList.add('d-none');
                infoError.classList.add('d-none');
                
                // Fetch file information
                fetch(`/api/info/${filename}`)
                    .then(response => {
                        if (!response.ok) {
                            throw new Error('Failed to load file information');
                        }
                        return response.json();
                    })
                    .then(data => {
                        // Fill in the info fields
                        document.getElementById('info-filename').textContent = data.name || filename;
                        document.getElementById('info-size').textContent = data.size_formatted || 'Unknown';
                        document.getElementById('info-modified').textContent = data.modified || 'Unknown';
                        
                        // Excel specific details
                        document.getElementById('info-sheets').textContent = 
                            Array.isArray(data.sheets) ? data.sheets.join(', ') : 'Unknown';
                        document.getElementById('info-active-sheet').textContent = data.active_sheet || 'Unknown';
                        
                        // Properties (author, title, etc)
                        let propertiesHtml = '';
                        if (data.properties && typeof data.properties === 'object') {
                            for (const [key, value] of Object.entries(data.properties)) {
                                if (value) {
                                    propertiesHtml += `<div><strong>${key}:</strong> ${value}</div>`;
                                }
                            }
                        }
                        document.getElementById('info-properties').innerHTML = 
                            propertiesHtml || 'No properties available';
                        
                        // Show content, hide loading and error
                        infoLoading.classList.add('d-none');
                        infoContent.classList.remove('d-none');
                    })
                    .catch(error => {
                        // Show error, hide loading and content
                        infoLoading.classList.add('d-none');
                        infoContent.classList.add('d-none');
                        infoError.classList.remove('d-none');
                        infoError.textContent = error.message;
                    });
            });
        }
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
        
        // Handle delete confirmation modal
        const deleteModal = document.getElementById('deleteModal');
        if (deleteModal) {
            deleteModal.addEventListener('show.bs.modal', function (event) {
                const button = event.relatedTarget;
                const filename = button.getAttribute('data-filename');
                const deleteFilename = document.getElementById('deleteFilename');
                const confirmDelete = document.getElementById('confirmDelete');
                
                if (deleteFilename && confirmDelete) {
                    deleteFilename.textContent = filename;
                    confirmDelete.href = `/delete/${filename}`;
                }
            });
        }
    }
    
    // Initial adjustments
    adjustFilesContainerHeight();
    setupDropdownPositioning();
    setupTableScrolling();
    setupFileUpload();
    setupInfoModal();
    setupModalHandlers();
    
    // Adjust on window resize
    window.addEventListener('resize', debounce(() => {
        adjustFilesContainerHeight();
        setupDropdownPositioning();
    }, 200));
    
    // Add a small delay to make sure everything is rendered properly
    setTimeout(() => {
        adjustFilesContainerHeight();
        setupDropdownPositioning();
    }, 200);
    
    // Utilities
    function debounce(func, wait) {
        let timeout;
        return function() {
            const context = this;
            const args = arguments;
            clearTimeout(timeout);
            timeout = setTimeout(() => func.apply(context, args), wait);
        };
    }
});
