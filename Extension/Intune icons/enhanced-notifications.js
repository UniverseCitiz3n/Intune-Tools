// Enhanced notification function to provide better positioning
function showEnhancedNotification(message, type = 'info', buttonElement = null) {
    // Remove any existing notification
    const existingNotification = document.getElementById('intune-tools-notification');
    if (existingNotification) {
        existingNotification.remove();
    }
    
    // Ensure styles are injected
    injectCustomStyles();
    
    // Create notification element
    const notification = document.createElement('div');
    notification.id = 'intune-tools-notification';
    notification.className = `intune-notification intune-notification-${type}`;
    
    // Create inner text container
    const textSpan = document.createElement('span');
    textSpan.textContent = message;
    notification.appendChild(textSpan);
    
    // Try to position near the button if available
    if (buttonElement && buttonElement.getBoundingClientRect) {
        const rect = buttonElement.getBoundingClientRect();
        const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
        
        // Position notification below the button
        notification.style.position = 'absolute';
        notification.style.top = `${rect.bottom + scrollTop + 2}px`;
        notification.style.left = `${rect.left}px`;
        notification.style.transform = 'none';
        notification.style.right = 'auto';
        
        // Add to body with absolute positioning
        document.body.appendChild(notification);
    } else {
        // Use the standard notification system
        let notificationContainer = document.querySelector('.ms-CommandBar') || 
                                  document.querySelector('div[role="menubar"]') ||
                                  document.querySelector('div[class*="ms-OverflowSet ms-CommandBar-primaryCommand"]');
        
        if (notificationContainer) {
            // Try to get the parent for better positioning
            const parent = notificationContainer.parentElement;
            if (parent) {
                parent.style.position = 'relative'; 
                parent.appendChild(notification);
            } else {
                notificationContainer.style.position = 'relative';
                notificationContainer.appendChild(notification);
            }
        } else {
            // Fallback if command bar not found, place in body with fixed position
            notification.style.position = 'fixed';
            notification.style.top = '15px';
            notification.style.right = '15px';
            notification.style.transform = 'none';
            document.body.appendChild(notification);
        }
    }
    
    // Auto-remove after a delay, unless it's an error
    if (type !== 'error') {
        setTimeout(() => {
            if (notification.parentNode) {
                notification.remove();
            }
        }, 3000);
    } else {
        // Add a close button for errors
        const closeBtn = document.createElement('span');
        closeBtn.className = 'intune-notification-close';
        closeBtn.innerHTML = '&times;';
        closeBtn.addEventListener('click', () => notification.remove());
        notification.appendChild(closeBtn);
    }
    
    return notification;
}
