// This file contains debug helper functions for the extension

// Function to log all command bars and their structures on the page
function debugCommandBars() {
    console.log('=========== Intune Tools Debug ===========');
    
    // First check for the exact Fluent UI command bar structure we need
    console.log('Checking for Fluent UI command bar...');
    const fluentCommandBar = document.querySelector('.ms-FocusZone.ms-CommandBar');
    if (fluentCommandBar) {
        console.log('✅ Found Fluent UI CommandBar:', fluentCommandBar);
        console.log('Command Bar role:', fluentCommandBar.getAttribute('role'));
        console.log('Command Bar aria-label:', fluentCommandBar.getAttribute('aria-label'));
        
        // Find the overflow set
        const overflowSet = fluentCommandBar.querySelector('.ms-OverflowSet');
        if (overflowSet) {
            console.log('✅ Found OverflowSet container:', overflowSet);
            console.log('OverflowSet role:', overflowSet.getAttribute('role'));
            console.log('OverflowSet class name:', overflowSet.className);
            
            // Find all overflow items (buttons)
            const overflowItems = overflowSet.querySelectorAll('.ms-OverflowSet-item');
            console.log(`Found ${overflowItems.length} overflow items`);
            
            Array.from(overflowItems).forEach((item, i) => {
                const button = item.querySelector('button');
                if (button) {
                    console.log(`  Button ${i}: aria-label=${button.getAttribute('aria-label')}, text=${button.textContent?.trim()}`);
                    console.log(`  Button ${i} class:`, button.className);
                }
            });
        }
    } else {
        console.log('❌ Could not find Fluent UI CommandBar');
    }
    
    // Look for common command bar selectors
    const selectors = [
        'div[role="menubar"]',
        'div[class*="CommandBar"]',
        'div[class*="OverflowSet"]',
        '.ms-CommandBar',
        '.ms-CommandBar-primaryCommand',
        '.ms-OverflowSet-item',
        'ul[role="toolbar"]'
    ];
    
    console.log('Searching for command bars using selectors:', selectors);
    
    // Find all elements that might be command bars
    const potentialElements = [];
    selectors.forEach(selector => {
        const elements = document.querySelectorAll(selector);
        if (elements.length > 0) {
            console.log(`Found ${elements.length} elements matching "${selector}":`);
            elements.forEach((el, i) => {
                console.log(`  [${i}] Element:`, el);
                
                // Check if this has buttons we care about
                const buttons = Array.from(el.querySelectorAll('button'));
                const buttonLabels = buttons.map(b => 
                    b.getAttribute('aria-label') || b.textContent?.trim() || 'unnamed button'
                );
                console.log(`    Contains buttons: ${buttonLabels.join(', ')}`);
                
                potentialElements.push(el);
            });
        }
    });
    
    // Look for specific buttons and trace up to find their command bars
    const buttonLabels = ['Refresh', 'Export', 'Columns'];
    console.log(`Looking for specific buttons: ${buttonLabels.join(', ')}`);
    
    buttonLabels.forEach(label => {
        // Find buttons by aria-label or text content
        const buttons = Array.from(document.querySelectorAll('button')).filter(btn => 
            btn.getAttribute('aria-label') === label || 
            btn.textContent?.trim() === label
        );
        
        if (buttons.length > 0) {
            console.log(`Found ${buttons.length} "${label}" buttons:`);
            buttons.forEach((btn, i) => {
                console.log(`  [${i}] Button:`, btn);
                
                // Trace up to potential command bar
                let parent = btn.parentElement;
                const path = [];
                while (parent && parent !== document.body) {
                    const desc = `${parent.tagName.toLowerCase()}${parent.id ? '#'+parent.id : ''}${
                        parent.className ? '.'+parent.className.replace(/\s+/g, '.') : ''
                    }[role=${parent.getAttribute('role') || 'none'}]`;
                    path.push(desc);
                    parent = parent.parentElement;
                }
                
                console.log(`    Parent hierarchy: ${path.join(' > ')}`);
            });
        } else {
            console.log(`No "${label}" buttons found on page`);
        }
    });
    
    console.log('=========== End Debug ===========');
    return potentialElements;
}

// Export for content script
window.debugCommandBars = debugCommandBars;
