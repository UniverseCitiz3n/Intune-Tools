# Intune Tools Extension

A browser extension that adds custom functionality to specific pages in the Microsoft Intune admin center.

## Features

- Adds a "Run Script" button to the toolbar on Intune device group membership pages
- Only appears on URLs matching the specific device group membership path
- Executes a series of fetch requests when the button is clicked
- Shows notifications for script execution status
- Debug mode to help identify UI elements (Ctrl+Click on the button)

## Installation

### Developer Mode Installation

1. Clone or download this repository
2. Open your browser (Chrome or Edge)
3. Go to the extensions page:
   - Chrome: `chrome://extensions`
   - Edge: `edge://extensions`
4. Enable "Developer mode" (toggle switch in top right)
5. Click "Load unpacked"
6. Select the folder containing this extension

### From Store (Future)

Once published, the extension can be installed from:
- Chrome Web Store (link TBD)
- Microsoft Edge Add-ons (link TBD)

## Usage

1. Navigate to the Intune admin center in your browser
2. Go to a device group membership page (URL pattern: `https://intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/groupMembership/`)
3. The "Run Script" button will appear in the command bar next to standard Intune buttons (like Refresh, Export, Columns)
4. Click the button to execute the script
5. A notification will appear showing the execution status

### Debug Mode

If you encounter issues with the button not appearing in the correct location:

1. Press and hold Ctrl while clicking the button (if visible)
2. This activates debug mode and logs information about potential command bars to the console
3. Open browser developer tools (F12) and check the console logs for diagnostic information

## Troubleshooting

If the button doesn't appear:
1. Ensure you're on the correct Intune page path
2. Refresh the page
3. Check the browser console for error messages
4. Try using debug mode (if button is visible) to diagnose command bar detection issues

## Development

The extension consists of:
- `manifest.json`: Extension configuration
- `content.js`: Main script that injects UI and handles functionality
- `debug.js`: Debugging utilities
- `enhanced-notifications.js`: Improved notification system
- `background.js`: Background script for extension events
- `icons/`: Extension icons

### Command Bar Detection

The extension uses multiple strategies to find the correct command bar:
1. Looks for specific buttons (Refresh, Export, Columns)
2. Navigates up the DOM to find their containing command bar
3. Uses DOM mutation observers to detect dynamic UI changes
4. Falls back to CSS selector-based detection methods

### Recent Changes

- Improved command bar detection for greater reliability
- Added debug mode to help diagnose UI issues
- Enhanced notification system with better positioning
- More robust error handling
- Added mutations observer to detect dynamic UI changes

1. Navigate to the Microsoft Intune admin center device group membership page
   (URL starting with `intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/groupMembership/`)
2. Look for the "Run Script" button in the toolbar (only appears on matching pages)
3. Click the button to execute the custom script

## Development

### Project Structure

- `manifest.json`: Extension configuration
- `content.js`: Script that injects the button and handles UI interactions
- `background.js`: Background script for handling extension events
- `icons/`: Directory containing extension icons

### Modifying the Script

To change what happens when the button is clicked, modify the `runFetchRequests()` function in `content.js`.

## License

[MIT License](LICENSE)

## Disclaimer

This extension is not affiliated with or endorsed by Microsoft. Use at your own risk.
