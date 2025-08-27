document.addEventListener("DOMContentLoaded", () => {
  // Global state variables
  const state = {
    currentDisplayType: 'config',
    sortDirection: 'asc',
    theme: 'light',
    targetMode: 'device', // New: track whether we're targeting devices or users
    selectedTableRows: new Set(), // Track selected table rows
    dynamicGroups: new Set(), // Track dynamic groups that cannot be modified manually
    pagination: {
      currentPage: 1,
      itemsPerPage: 10,
      totalItems: 0,
      totalPages: 0,
      filteredData: [],
      selectedRowIds: new Set() // Track selected rows across pages by unique identifier
    }
  };

  // ── Theme Management Functions ───────────────────────────────────────
  const toggleTheme = () => {
    state.theme = state.theme === 'light' ? 'dark' : 'light';
    applyTheme(state.theme);
    chrome.storage.local.set({ theme: state.theme });
    logMessage(`Theme switched to ${state.theme} mode`);
  };

  const applyTheme = (theme) => {
    document.body.setAttribute('data-theme', theme);
  };

  const initializeTheme = () => {
    chrome.storage.local.get('theme', (data) => {
      if (data.theme) {
        state.theme = data.theme;
        applyTheme(state.theme);
        logMessage(`Theme initialized to ${state.theme} mode`);
      }
    });
  };

  // ── Group Type Management Functions ───────────────────────────────────
  const isDynamicGroup = (groupId) => {
    return state.dynamicGroups.has(groupId);
  };

  const addDynamicGroup = (groupId) => {
    state.dynamicGroups.add(groupId);
  };
  const clearDynamicGroups = () => {
    state.dynamicGroups.clear();
  };

  // fetchGroupWithType: Fetch group details including groupTypes
  const fetchGroupWithType = async (groupId, token) => {
    try {
      const groupData = await fetchJSON(`https://graph.microsoft.com/v1.0/groups/${groupId}?$select=id,displayName,groupTypes`, {
        method: "GET",
        headers: {
          "Authorization": token,
          "Content-Type": "application/json"
        }
      });

      // Check if group is dynamic
      const isDynamic = groupData.groupTypes && groupData.groupTypes.includes('DynamicMembership');
      if (isDynamic) {
        addDynamicGroup(groupData.id);
      }

      return {
        id: groupData.id,
        displayName: groupData.displayName,
        isDynamic: isDynamic
      };
    } catch (error) {
      logMessage(`Error fetching group details for ${groupId}: ${error.message}`);
      return null;
    }
  };

  // ── Utility Functions ───────────────────────────────────────────────
  // filterTable: Filter the table based on input text
  const filterTable = (filterText) => {
    filterText = filterText.toLowerCase();
    if (state.currentDisplayType === 'config') {
      chrome.storage.local.get(['lastConfigAssignments'], (data) => {
        if (data.lastConfigAssignments) {
          const filteredResults = [...data.lastConfigAssignments].filter(item =>
            item.policyName.toLowerCase().includes(filterText)
          );

          updateConfigTable(filteredResults, false); // Pass false to avoid changing currentDisplayType
        }
      });
    } else if (state.currentDisplayType === 'apps') {
      chrome.storage.local.get(['lastAppAssignments'], (data) => {
        if (data.lastAppAssignments) {
          const filteredResults = [...data.lastAppAssignments].filter(item =>
            item.appName.toLowerCase().includes(filterText)
          );

          updateAppTable(filteredResults, false); // Pass false to avoid changing currentDisplayType
        }
      });
    } else if (state.currentDisplayType === 'compliance') {
      chrome.storage.local.get(['lastComplianceAssignments'], (data) => {
        if (data.lastComplianceAssignments) {
          const filteredResults = [...data.lastComplianceAssignments].filter(item =>
            item.policyName.toLowerCase().includes(filterText)
          );

          updateComplianceTable(filteredResults, false); // Pass false to avoid changing currentDisplayType
        }
      });
    } else if (state.currentDisplayType === 'pwsh') {
      chrome.storage.local.get(['lastPwshAssignments'], (data) => {
        if (data.lastPwshAssignments) {
          const filteredResults = [...data.lastPwshAssignments].filter(item =>
            item.scriptName.toLowerCase().includes(filterText)
          );

          updatePwshTable(filteredResults, false); // Pass false to avoid changing currentDisplayType
        }
      });
    } else if (state.currentDisplayType === 'groupMembers') {
      chrome.storage.local.get(['lastGroupMembers'], (data) => {
        if (data.lastGroupMembers) {
          const filteredResults = [...data.lastGroupMembers].filter(item => {
            const name = (item.displayName || '').toLowerCase();
            const upn = (item.userPrincipalName || '').toLowerCase();
            const deviceId = (item.deviceId || '').toLowerCase();
            return name.includes(filterText) || upn.includes(filterText) || deviceId.includes(filterText);
          });

          updateGroupMembersTable(filteredResults, false);
        }
      });
    }
  };

  // ── Pagination Functions ────────────────────────────────────────────
  // updatePaginationState: Update pagination state with filtered data
  const updatePaginationState = (data, filterText = '') => {
    const filteredData = filterText ? 
      data.filter(item => {
        const searchText = filterText.toLowerCase();
        if (state.currentDisplayType === 'config') {
          return item.policyName.toLowerCase().includes(searchText);
        } else if (state.currentDisplayType === 'apps') {
          return item.appName.toLowerCase().includes(searchText);
        } else if (state.currentDisplayType === 'compliance') {
          return item.policyName.toLowerCase().includes(searchText);
        } else if (state.currentDisplayType === 'pwsh') {
          return item.scriptName.toLowerCase().includes(searchText);
        } else if (state.currentDisplayType === 'groupMembers') {
          const name = (item.displayName || '').toLowerCase();
          const upn = (item.userPrincipalName || '').toLowerCase();
          const deviceId = (item.deviceId || '').toLowerCase();
          return name.includes(searchText) || upn.includes(searchText) || deviceId.includes(searchText);
        }
        return true;
      }) : data;

    state.pagination.filteredData = filteredData;
    state.pagination.totalItems = filteredData.length;
    state.pagination.totalPages = Math.ceil(filteredData.length / state.pagination.itemsPerPage);
    
    // Reset to page 1 if current page is beyond available pages
    if (state.pagination.currentPage > state.pagination.totalPages) {
      state.pagination.currentPage = Math.max(1, state.pagination.totalPages);
    }
  };

  // getCurrentPageData: Get data for current page
  const getCurrentPageData = () => {
    const startIndex = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage;
    const endIndex = startIndex + state.pagination.itemsPerPage;
    return state.pagination.filteredData.slice(startIndex, endIndex);
  };

  // generatePageNumbers: Generate smart page number display
  const generatePageNumbers = () => {
    const { currentPage, totalPages } = state.pagination;
    const pages = [];
    
    if (totalPages <= 7) {
      // Show all pages if 7 or fewer
      for (let i = 1; i <= totalPages; i++) {
        pages.push(i);
      }
    } else {
      // Smart pagination with ellipsis
      if (currentPage <= 4) {
        // Near beginning: 1 2 3 4 5 ... last
        for (let i = 1; i <= 5; i++) {
          pages.push(i);
        }
        pages.push('...');
        pages.push(totalPages);
      } else if (currentPage >= totalPages - 3) {
        // Near end: 1 ... last-4 last-3 last-2 last-1 last
        pages.push(1);
        pages.push('...');
        for (let i = totalPages - 4; i <= totalPages; i++) {
          pages.push(i);
        }
      } else {
        // Middle: 1 ... current-1 current current+1 ... last
        pages.push(1);
        pages.push('...');
        for (let i = currentPage - 1; i <= currentPage + 1; i++) {
          pages.push(i);
        }
        pages.push('...');
        pages.push(totalPages);
      }
    }
    
    return pages;
  };

  // updatePaginationControls: Update pagination UI elements
  const updatePaginationControls = () => {
    const container = document.getElementById('paginationContainer');
    const resultsSpan = document.getElementById('paginationResults');
    const pageNumbers = document.getElementById('pageNumbers');
    const prevBtn = document.getElementById('prevPageBtn');
    const nextBtn = document.getElementById('nextPageBtn');

    if (state.pagination.totalItems === 0) {
      container.style.display = 'none';
      return;
    }

    container.style.display = 'flex';

    // Update results info
    const startItem = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage + 1;
    const endItem = Math.min(state.pagination.currentPage * state.pagination.itemsPerPage, state.pagination.totalItems);
    resultsSpan.textContent = `Showing ${startItem}–${endItem} of ${state.pagination.totalItems} results`;

    // Update page numbers
    const pages = generatePageNumbers();
    pageNumbers.innerHTML = '';
    
    pages.forEach(page => {
      if (page === '...') {
        const ellipsis = document.createElement('span');
        ellipsis.className = 'page-ellipsis';
        ellipsis.textContent = '...';
        pageNumbers.appendChild(ellipsis);
      } else {
        const pageBtn = document.createElement('button');
        pageBtn.className = 'page-number';
        pageBtn.textContent = page;
        if (page === state.pagination.currentPage) {
          pageBtn.classList.add('active');
        }
        pageBtn.addEventListener('click', () => goToPage(page));
        pageNumbers.appendChild(pageBtn);
      }
    });

    // Update prev/next buttons
    prevBtn.disabled = state.pagination.currentPage === 1;
    nextBtn.disabled = state.pagination.currentPage === state.pagination.totalPages;
  };

  // clearTableAndPagination: Helper function to clear table and hide pagination
  const clearTableAndPagination = () => {
    document.getElementById("configTableBody").innerHTML = '';
    document.getElementById('paginationContainer').style.display = 'none';
    state.pagination.totalItems = 0;
    state.pagination.totalPages = 0;
    state.pagination.currentPage = 1;
    state.pagination.filteredData = [];
    state.pagination.selectedRowIds.clear(); // Clear selection when clearing table
  };

  // generateRowId: Generate unique ID for a table row based on content
  const generateRowId = (rowData) => {
    if (state.currentDisplayType === 'config') {
      return `${rowData.policyName}-${rowData.targets[0].groupName}-${rowData.targets[0].targetType}`;
    } else if (state.currentDisplayType === 'apps') {
      return `${rowData.appName}-${rowData.targets[0].groupName}-${rowData.targets[0].targetType}`;
    } else if (state.currentDisplayType === 'compliance') {
      return `${rowData.policyName}-${rowData.targets[0].groupName}-${rowData.targets[0].targetType}`;
    } else if (state.currentDisplayType === 'pwsh') {
      return `${rowData.scriptName}-${rowData.targetName}`;
    }
    return '';
  };

  // restoreRowSelection: Restore row selection after rendering
  const restoreRowSelection = () => {
    document.querySelectorAll('#configTableBody tr').forEach((row) => {
      const rowId = row.dataset.rowId;
      if (rowId && state.pagination.selectedRowIds.has(rowId)) {
        row.classList.add('table-row-selected');
      }
    });
  };

  // clearTableSelection: Clear all table row selections
  const clearTableSelection = () => {
    document.querySelectorAll('.table-row-selected').forEach(row => {
      row.classList.remove('table-row-selected');
    });
    state.pagination.selectedRowIds.clear();
  };

  // goToPage: Navigate to specific page
  const goToPage = (page) => {
    if (page < 1 || page > state.pagination.totalPages) return;
    
    state.pagination.currentPage = page;
    renderCurrentPage();
    updatePaginationControls();
  };

  // renderCurrentPage: Render current page data in table
  const renderCurrentPage = () => {
    const currentPageData = getCurrentPageData();
    
    if (state.currentDisplayType === 'config') {
      renderConfigTablePage(currentPageData);
    } else if (state.currentDisplayType === 'apps') {
      renderAppTablePage(currentPageData);
    } else if (state.currentDisplayType === 'compliance') {
      renderComplianceTablePage(currentPageData);
    } else if (state.currentDisplayType === 'pwsh') {
      renderPwshTablePage(currentPageData);
    } else if (state.currentDisplayType === 'groupMembers') {
      renderGroupMembersTablePage(currentPageData);
    }
  };

  // Individual table rendering functions
  const renderConfigTablePage = (assignments) => {
    let rows = '';
    let rowIndex = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage;
    
    assignments.forEach(policy => {
      policy.targets.forEach(target => {
        // Generate unique row ID
        const rowData = { policyName: policy.policyName, targets: [target] };
        const rowId = generateRowId(rowData);
        
        // Check if this is a virtual assignment or dynamic group
        const isVirtualAssignment = target.groupName === 'All Devices' || target.groupName === 'All Users';
        const isDynamicGroupAssignment = target.groupId && isDynamicGroup(target.groupId);
        const isDisabled = isVirtualAssignment || isDynamicGroupAssignment;
        const disabledClass = isDisabled ? 'table-row-disabled' : 'table-row-selectable';

        // Add tooltip for disabled rows
        let tooltipText = '';
        if (isVirtualAssignment) {
          tooltipText = ' title="Virtual group"';
        } else if (isDynamicGroupAssignment) {
          tooltipText = ' title="Dynamic group – cannot modify manually"';
        }

        rows += `<tr class="${disabledClass}" data-row-index="${rowIndex}" data-row-id="${rowId}"${tooltipText}>
          <td style="word-wrap: break-word; white-space: normal;">${policy.policyName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.groupName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.membershipType}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.targetType}</td>
        </tr>`;
        rowIndex++;
      });
    });

    document.getElementById("configTableBody").innerHTML = rows;

    // Add event listeners for row clicks
    document.querySelectorAll('#configTableBody tr').forEach((row, index) => {
      row.addEventListener('click', (e) => handleTableRowClick(row, index));
    });

    // Restore selection state
    restoreRowSelection();
  };

  const renderAppTablePage = (assignments) => {
    let rows = '';
    let rowIndex = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage;
    
    assignments.forEach(app => {
      app.targets.forEach(target => {
        // Generate unique row ID
        const rowData = { appName: app.appName, targets: [target] };
        const rowId = generateRowId(rowData);
        
        // Check if this is a virtual assignment or dynamic group
        const isVirtualAssignment = target.groupName === 'All Devices' || target.groupName === 'All Users';
        const isDynamicGroupAssignment = target.groupId && isDynamicGroup(target.groupId);
        const isDisabled = isVirtualAssignment || isDynamicGroupAssignment;
        const disabledClass = isDisabled ? 'table-row-disabled' : 'table-row-selectable';

        // Add tooltip for disabled rows
        let tooltipText = '';
        if (isVirtualAssignment) {
          tooltipText = ' title="Virtual group"';
        } else if (isDynamicGroupAssignment) {
          tooltipText = ' title="Dynamic group – cannot modify manually"';
        }

        rows += `<tr class="${disabledClass}" data-row-index="${rowIndex}" data-row-id="${rowId}"${tooltipText}>
          <td style="word-wrap: break-word; white-space: normal;">${app.appName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.groupName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.membershipType}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.targetType}</td>
        </tr>`;
        rowIndex++;
      });
    });

    document.getElementById("configTableBody").innerHTML = rows;

    // Add event listeners for row clicks
    document.querySelectorAll('#configTableBody tr').forEach((row, index) => {
      row.addEventListener('click', (e) => handleTableRowClick(row, index));
    });

    // Restore selection state
    restoreRowSelection();
  };

  const renderComplianceTablePage = (assignments) => {
    let rows = '';
    let rowIndex = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage;
    
    assignments.forEach(policy => {
      policy.targets.forEach(target => {
        // Generate unique row ID
        const rowData = { policyName: policy.policyName, targets: [target] };
        const rowId = generateRowId(rowData);
        
        // Check if this is a virtual assignment or dynamic group
        const isVirtualAssignment = target.groupName === 'All Devices' || target.groupName === 'All Users';
        const isDynamicGroupAssignment = target.groupId && isDynamicGroup(target.groupId);
        const isDisabled = isVirtualAssignment || isDynamicGroupAssignment;
        const disabledClass = isDisabled ? 'table-row-disabled' : 'table-row-selectable';

        // Add tooltip for disabled rows
        let tooltipText = '';
        if (isVirtualAssignment) {
          tooltipText = ' title="Virtual group"';
        } else if (isDynamicGroupAssignment) {
          tooltipText = ' title="Dynamic group – cannot modify manually"';
        }

        rows += `<tr class="${disabledClass}" data-row-index="${rowIndex}" data-row-id="${rowId}"${tooltipText}>
          <td style="word-wrap: break-word; white-space: normal;">${policy.policyName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.groupName}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.membershipType}</td>
          <td style="word-wrap: break-word; white-space: normal;">${target.targetType}</td>
        </tr>`;
        rowIndex++;
      });
    });

    document.getElementById("configTableBody").innerHTML = rows;

    // Add event listeners for row clicks
    document.querySelectorAll('#configTableBody tr').forEach((row, index) => {
      row.addEventListener('click', (e) => handleTableRowClick(row, index));
    });

    // Restore selection state
    restoreRowSelection();
  };

  const renderPwshTablePage = (assignments) => {
    let rows = '';
    let rowIndex = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage;
    
    assignments.forEach(script => {
      // Generate unique row ID
      const rowData = { scriptName: script.scriptName, targetName: script.targetName };
      const rowId = generateRowId(rowData);
      
      // Check if this is a virtual assignment or dynamic group
      const isVirtualAssignment = script.targetName === 'All Devices' || script.targetName === 'All Users';
      const isDynamicGroupAssignment = script.targetGroupId && isDynamicGroup(script.targetGroupId);
      const isDisabled = isVirtualAssignment || isDynamicGroupAssignment;
      const disabledClass = isDisabled ? 'table-row-disabled' : 'table-row-selectable';

      // Add tooltip for disabled rows
      let tooltipText = '';
      if (isVirtualAssignment) {
        tooltipText = ' title="Virtual group"';
      } else if (isDynamicGroupAssignment) {
        tooltipText = ' title="Dynamic group – cannot modify manually"';
      }

      rows += `<tr class="${disabledClass}" data-row-index="${rowIndex}" data-row-id="${rowId}"${tooltipText}>
        <td style="word-wrap: break-word; white-space: normal;">${script.scriptName}</td>
        <td style="word-wrap: break-word; white-space: normal;">${script.description || ''}</td>
        <td style="word-wrap: break-word; white-space: normal;">${script.targetName}</td>
      </tr>`;
      rowIndex++;
    });

    document.getElementById("configTableBody").innerHTML = rows;

    // Add event listeners for row clicks
    document.querySelectorAll('#configTableBody tr').forEach((row, index) => {
      row.addEventListener('click', (e) => handleTableRowClick(row, index));
    });

    // Restore selection state
    restoreRowSelection();
  };

  const renderGroupMembersTablePage = (members) => {
    let rows = '';
    let rowIndex = (state.pagination.currentPage - 1) * state.pagination.itemsPerPage;

    members.forEach(member => {
      const objectType = (member['@odata.type'] || '').split('.').pop();
      const upnOrId = member.userPrincipalName || member.deviceId || '';
      rows += `<tr data-row-index="${rowIndex}">
        <td style="word-wrap: break-word; white-space: normal;">${member.displayName || ''}</td>
        <td style="word-wrap: break-word; white-space: normal;">${upnOrId}</td>
        <td style="word-wrap: break-word; white-space: normal;">${objectType}</td>
      </tr>`;
      rowIndex++;
    });

    document.getElementById("configTableBody").innerHTML = rows;
  };

  // fetchJSON: Helper to fetch and parse JSON responses
  const fetchJSON = async (url, options = {}) => {
    const response = await fetch(url, options);
    return parseJSON(response);
  };

  // verifyMdmUrl: Verify the current tab URL and extract mdmDeviceId
  const verifyMdmUrl = async () => {
    return new Promise((resolve, reject) => {
      chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
        if (!tabs || !tabs[0]) {
          const error = 'No active tab found.';
          logMessage(error);
          showNotification(error, 'error');
          return reject(new Error(error));
        }
        const url = tabs[0].url;
        const sanitizedUrl = url.replace(/\/[\w-]+(?=\/|$)/g, '/[REDACTED]');
        logMessage(`Active tab URL: ${sanitizedUrl}`);
        const mdmMatch = url.match(/mdmDeviceId\/([\w-]+)/i);
        if (!mdmMatch) {
          const error = 'mdmDeviceId not found in URL.';
          logMessage(error);
          showNotification(error, 'error');
          return reject(new Error(error));
        }
        resolve({ mdmDeviceId: mdmMatch[1] });
      });
    });
  };

  // getToken: Retrieve token from Chrome storage
  const getToken = async () => {
    return new Promise((resolve, reject) => {
      chrome.storage.local.get("msGraphToken", (data) => {
        if (data.msGraphToken) {
          resolve(data.msGraphToken);
        } else {
          const error = 'No token captured. Please login first.';
          logMessage(error);
          showNotification(error, 'error');
          reject(new Error(error));
        }
      });
    });  
  };  // getAllGroupsMap: Get groups map (device & user) for lookups, with dynamic group tracking
  const getAllGroupsMap = async (deviceObjectId, userObjectId, token) => {
    const headers = {
      "Authorization": token,
      "Content-Type": "application/json",
      "ConsistencyLevel": "eventual"
    };

    // Endpoints to get group memberships with groupTypes
    const endpoints = [
      `https://graph.microsoft.com/beta/devices/${deviceObjectId}/memberOf?$select=id,displayName,groupTypes&$orderBy=displayName%20asc&$count=true`,
      `https://graph.microsoft.com/beta/devices/${deviceObjectId}/transitiveMemberOf?$select=id,displayName,groupTypes&$orderBy=displayName%20asc&$count=true`
    ];

    if (userObjectId) {
      endpoints.push(
        `https://graph.microsoft.com/beta/users/${userObjectId}/memberOf?$select=id,displayName,groupTypes&$orderBy=displayName%20asc&$count=true`,
        `https://graph.microsoft.com/beta/users/${userObjectId}/transitiveMemberOf?$select=id,displayName,groupTypes&$orderBy=displayName%20asc&$count=true`
      );
    }

    const results = await Promise.all(
      endpoints.map(url => fetchJSON(url, { method: "GET", headers }))
    );

    const groupMap = new Map();

    results.forEach(result => {
      (result.value || []).forEach(group => {
        if (group['@odata.type'] === '#microsoft.graph.group') {
          groupMap.set(group.id, group.displayName);

          // Track dynamic groups
          const isDynamic = group.groupTypes && group.groupTypes.includes('DynamicMembership');
          if (isDynamic) {
            addDynamicGroup(group.id);
          }
        }
      });
    });

    return groupMap;
  };

  // getDirectoryObjectId: Get directory object ID for device or user based on current mode
  const getDirectoryObjectId = async (mdmDeviceId, token, mode = state.targetMode) => {
    logMessage(`getDirectoryObjectId: Getting ${mode} directory object ID`);

    // First, get the managed device data
    const managedDeviceData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/manageddevices('${mdmDeviceId}')`, {
      method: "GET",
      headers: { "Authorization": token, "Content-Type": "application/json" }
    });

    if (mode === 'device') {
      const azureADDeviceId = managedDeviceData.azureADDeviceId;
      if (!azureADDeviceId) throw new Error("azureADDeviceId not found in managed device data.");

      const filter = encodeURIComponent(`deviceId eq '${azureADDeviceId}'`);
      const deviceData = await fetchJSON(`https://graph.microsoft.com/v1.0/devices?$top=100&$filter=${filter}`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      if (!deviceData.value || deviceData.value.length === 0) {
        throw new Error("No device found matching the azureADDeviceId.");
      }

      return {
        directoryId: deviceData.value[0].id,
        displayName: deviceData.value[0].displayName || 'Unknown Device'
      };
    } else if (mode === 'user') {
      const userPrincipalName = managedDeviceData.userPrincipalName;
      if (!userPrincipalName || userPrincipalName === 'Unknown user') {
        throw new Error("No user associated with this device.");
      }

      const userData = await fetchJSON(`https://graph.microsoft.com/beta/users?$filter=userPrincipalName eq '${encodeURIComponent(userPrincipalName)}'`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      if (!userData.value || userData.value.length === 0) {
        throw new Error("No user found matching the userPrincipalName.");
      }

      return {
        directoryId: userData.value[0].id,
        displayName: userData.value[0].displayName || userPrincipalName
      };
    }

    throw new Error(`Invalid mode: ${mode}`);
  };  // ── UI Update Functions ──────────────────────────────────────────────  
  // updateDeviceNameDisplay: Update the device name display next to "Configuration Assignments"
  const updateDeviceNameDisplay = (deviceData) => {
    const deviceNameElement = document.getElementById('deviceNameDisplay');
    if (deviceNameElement && deviceData && deviceData.deviceName) {
      deviceNameElement.textContent = `- ${deviceData.deviceName}`;
      logMessage(`updateDeviceNameDisplay: Updated device name display to "${deviceData.deviceName}"`);
    }
  };
  // Table row selection helpers
  // (clearTableSelection moved to pagination section)

  // Clear checkbox selections
  const clearCheckboxSelection = () => {
    document.querySelectorAll("#groupResults input[type=checkbox]:checked").forEach(cb => {
      cb.checked = false;
    });
    // Update stored search results to reflect unchecked state
    chrome.storage.local.get(['lastSearchResults'], (data) => {
      if (data.lastSearchResults) {
        const updated = data.lastSearchResults.map(group => ({ ...group, checked: false }));
        chrome.storage.local.set({ lastSearchResults: updated });
      }
    });
    updateActionButtonsState();
  };

  const updateActionButtonsState = () => {
    const selected = document.querySelectorAll('#groupResults input[type=checkbox]:checked');
    const hasDynamic = Array.from(selected).some(cb => isDynamicGroup(cb.value));
    ['addToGroups', 'removeFromGroups'].forEach(id => {
      const btn = document.getElementById(id);
      if (!btn) return;
      if (hasDynamic) {
        btn.classList.add('disabled');
      } else {
        btn.classList.remove('disabled');
      }
    });
  };

  const getSelectedGroupNames = () => {
    const selectedGroups = [];
    document.querySelectorAll('.table-row-selected').forEach(row => {
      let groupNameCell;

      // Determine which column contains the group name based on current table type
      if (state.currentDisplayType === 'pwsh') {
        groupNameCell = row.children[2]; // Assignment Target is the 3rd column (index 2) for PowerShell scripts
      } else {
        groupNameCell = row.children[1]; // Group Name is the 2nd column (index 1) for other tables after removing checkbox column
      }

      const groupName = groupNameCell.textContent.trim();

      // Only include actual group names, not virtual assignments
      if (groupName && groupName !== 'All Devices' && groupName !== 'All Users') {
        selectedGroups.push(groupName);
      }
    });
    return [...new Set(selectedGroups)]; // Remove duplicates
  };
  const getAllSelectedGroups = () => {
    // Get groups from search results (exclude dynamic groups from operations)
    const searchResultGroups = [];
    const searchCheckboxes = document.querySelectorAll("#groupResults input[type=checkbox]:checked");
    searchCheckboxes.forEach(cb => {
      // Only include non-dynamic groups
      if (!isDynamicGroup(cb.value)) {
        searchResultGroups.push({
          id: cb.value,
          name: cb.dataset.groupName
        });
      }
    });

    // Get group names from selected table rows (new functionality)
    // Note: Dynamic groups are already filtered out by disabling row selection in table rendering
    const tableSelectedGroupNames = getSelectedGroupNames();

    return {
      searchResults: searchResultGroups,
      tableSelections: tableSelectedGroupNames,
      hasAnySelection: searchResultGroups.length > 0 || tableSelectedGroupNames.length > 0
    };
  }; const handleTableRowClick = (row, rowIndex) => {
    // Don't allow selection of disabled rows
    if (row.classList.contains('table-row-disabled')) {
      return;
    }

    // Clear checkbox selections when selecting table rows
    clearCheckboxSelection();

    const rowId = row.dataset.rowId;
    if (!rowId) return; // Skip if no row ID

    if (row.classList.contains('table-row-selected')) {
      // Deselect the row
      state.pagination.selectedRowIds.delete(rowId);
      row.classList.remove('table-row-selected');
    } else {
      // Select the row
      state.pagination.selectedRowIds.add(rowId);
      row.classList.add('table-row-selected');
    }
  };

  // updateButtonText: Update button text based on current target mode
  const updateButtonText = () => {
    const targetType = state.targetMode === 'device' ? 'Device' : 'User';
    document.getElementById('addBtnText').textContent = `Add ${targetType} to Groups`;
    document.getElementById('removeBtnText').textContent = `Remove ${targetType} from Groups`;
    logMessage(`updateButtonText: Updated buttons for ${targetType} mode`);
  };

  // handleTargetModeToggle: Handle switching between device and user modes
  const handleTargetModeToggle = (mode) => {
    logMessage(`handleTargetModeToggle: Called with mode '${mode}', current state is '${state.targetMode}'`);
    
    if (state.targetMode === mode) {
      logMessage(`handleTargetModeToggle: No change needed - already in ${mode} mode`);
      return; // No change needed
    }

    state.targetMode = mode;
    logMessage(`handleTargetModeToggle: State updated to '${mode}'`);

    // Update UI
    document.querySelectorAll('.target-type-toggle button').forEach(btn => {
      btn.classList.remove('active');
    });
    document.querySelector(`[data-mode="${mode}"]`).classList.add('active');
    logMessage(`handleTargetModeToggle: UI updated - '${mode}' button is now active`);

    updateButtonText();

    // Save to storage
    chrome.storage.local.set({ targetMode: state.targetMode });

    logMessage(`handleTargetModeToggle: Switched to ${mode} mode`);
  };  // updateTableHeaders: Update table header based on content type
  const updateTableHeaders = (type) => {
    const headerRow = document.querySelector('thead tr');
    let headerContent = '';
    if (type === 'config') {
      headerContent = `
        <th class="sortable" style="word-wrap: break-word; white-space: normal;">Profile Name</th>
        <th style="word-wrap: break-word; white-space: normal;">Group</th>
        <th style="word-wrap: break-word; white-space: normal;">Membership Type</th>
        <th style="word-wrap: break-word; white-space: normal;">Target Type</th>
      `;
    } else if (type === 'apps') {
      headerContent = `
        <th class="sortable" style="word-wrap: break-word; white-space: normal;">App Name</th>
        <th style="word-wrap: break-word; white-space: normal;">Group</th>
        <th style="word-wrap: break-word; white-space: normal;">Membership Type</th>
        <th style="word-wrap: break-word; white-space: normal;">Target Type</th>
      `;
    } else if (type === 'compliance') {
      headerContent = `
        <th class="sortable" style="word-wrap: break-word; white-space: normal;">Policy Name</th>
        <th style="word-wrap: break-word; white-space: normal;">Group</th>
        <th style="word-wrap: break-word; white-space: normal;">Membership Type</th>
        <th style="word-wrap: break-word; white-space: normal;">Target Type</th>
      `;
    } else if (type === 'pwsh') {
      headerContent = `
        <th class="sortable" style="word-wrap: break-word; white-space: normal;">Script Name</th>
        <th style="word-wrap: break-word; white-space: normal;">Description</th>
        <th style="word-wrap: break-word; white-space: normal;">Assignment Target</th>      `;
    } else if (type === 'groupMembers') {
      headerContent = `
        <th class="sortable" style="word-wrap: break-word; white-space: normal;">Display Name</th>
        <th style="word-wrap: break-word; white-space: normal;">UPN / Device ID</th>
        <th style="word-wrap: break-word; white-space: normal;">Object Type</th>
      `;
    }
    headerRow.innerHTML = headerContent;
    const sortableHeader = document.querySelector('th.sortable');
    if (sortableHeader) {
      sortableHeader.classList.remove('desc');
      sortableHeader.classList.add('asc');
    }
  };  // updateConfigTable: Update configuration assignments table
  const updateConfigTable = (assignments, updateDisplay = true) => {
    if (updateDisplay) {
      state.currentDisplayType = 'config';
      chrome.storage.local.set({ currentDisplayType: state.currentDisplayType });
    }
    updateTableHeaders('config');
    assignments.sort((a, b) => a.policyName.localeCompare(b.policyName));
    if (state.sortDirection === 'desc') assignments.reverse();

    // Flatten assignments for pagination (each target becomes a row)
    const flattenedData = [];
    assignments.forEach(policy => {
      policy.targets.forEach(target => {
        flattenedData.push({
          policyName: policy.policyName,
          targets: [target] // Keep single target for rendering
        });
      });
    });

    // Update pagination state
    const filterValue = document.getElementById('profileFilterInput').value.toLowerCase();
    updatePaginationState(flattenedData, filterValue);
    
    // Render current page
    renderCurrentPage();
    updatePaginationControls();

    const sortableHeader = document.querySelector('th.sortable');
    sortableHeader.classList.remove('desc');
    sortableHeader.classList.add('asc');
  };

  const updateGroupMembersTable = (members, updateDisplay = true) => {
    if (updateDisplay) {
      state.currentDisplayType = 'groupMembers';
      chrome.storage.local.set({ currentDisplayType: state.currentDisplayType });
    }
    updateTableHeaders('groupMembers');
    members.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));
    if (state.sortDirection === 'desc') members.reverse();

    const displayMembers = members.length > 100 ? members.slice(0, 100) : members;

    const flattenedData = displayMembers.map(m => ({
      displayName: m.displayName || '',
      userPrincipalName: m.userPrincipalName || '',
      deviceId: m.deviceId || '',
      ['@odata.type']: m['@odata.type'] || ''
    }));

    updatePaginationState(flattenedData);

    renderCurrentPage();
    updatePaginationControls();

    const sortableHeader = document.querySelector('th.sortable');
    if (sortableHeader) {
      sortableHeader.classList.remove('desc');
      sortableHeader.classList.add('asc');
    }
  };
  // updateAppTable: Update app assignments table (similar in structure)
  const updateAppTable = (assignments, updateDisplay = true) => {
    if (updateDisplay) {
      state.currentDisplayType = 'apps';
      chrome.storage.local.set({ currentDisplayType: state.currentDisplayType });
    }
    updateTableHeaders('apps');
    assignments.sort((a, b) => a.appName.localeCompare(b.appName));
    if (state.sortDirection === 'desc') assignments.reverse();

    // Flatten assignments for pagination (each target becomes a row)
    const flattenedData = [];
    assignments.forEach(app => {
      const targets = app.targets.filter(t => t.membershipType !== 'Not Member');
      targets.forEach(target => {
        flattenedData.push({
          appName: app.appName,
          appVersion: app.appVersion,
          installState: app.installState,
          targets: [target] // Keep single target for rendering
        });
      });
    });

    // Update pagination state
    const filterValue = document.getElementById('profileFilterInput').value.toLowerCase();
    updatePaginationState(flattenedData, filterValue);
    
    // Render current page
    renderCurrentPage();
    updatePaginationControls();

    const sortableHeader = document.querySelector('th.sortable');
    sortableHeader.classList.remove('desc');
    sortableHeader.classList.add('asc');
  };
  // updateComplianceTable: Update compliance assignments table
  const updateComplianceTable = (assignments, updateDisplay = true) => {
    if (updateDisplay) {
      state.currentDisplayType = 'compliance';
      chrome.storage.local.set({ currentDisplayType: state.currentDisplayType });
    }
    updateTableHeaders('compliance');
    assignments.sort((a, b) => a.policyName.localeCompare(b.policyName));
    if (state.sortDirection === 'desc') assignments.reverse();

    // Flatten assignments for pagination (each target becomes a row)
    const flattenedData = [];
    assignments.forEach(policy => {
      policy.targets.forEach(target => {
        flattenedData.push({
          policyName: policy.policyName,
          complianceState: policy.complianceState,
          targets: [target] // Keep single target for rendering
        });
      });
    });

    // Update pagination state
    const filterValue = document.getElementById('profileFilterInput').value.toLowerCase();
    updatePaginationState(flattenedData, filterValue);
    
    // Render current page
    renderCurrentPage();
    updatePaginationControls();

    const sortableHeader = document.querySelector('th.sortable');
    sortableHeader.classList.remove('desc');
    sortableHeader.classList.add('asc');
  };
  // updatePwshTable: Update PowerShell scripts table
  const updatePwshTable = (assignments, updateDisplay = true) => {
    if (updateDisplay) {
      state.currentDisplayType = 'pwsh';
      chrome.storage.local.set({ currentDisplayType: state.currentDisplayType });
    }
    updateTableHeaders('pwsh');
    assignments.sort((a, b) => a.scriptName.localeCompare(b.scriptName));
    if (state.sortDirection === 'desc') assignments.reverse();

    // For PowerShell scripts, each script is a single row (no targets to flatten)
    const flattenedData = assignments.map(script => ({
      scriptName: script.scriptName,
      description: script.description || '',
      targetName: script.targetName
    }));

    // Update pagination state
    const filterValue = document.getElementById('profileFilterInput').value.toLowerCase();
    updatePaginationState(flattenedData, filterValue);
    
    // Render current page
    renderCurrentPage();
    updatePaginationControls();

    const sortableHeader = document.querySelector('th.sortable');
    sortableHeader.classList.remove('desc');
    sortableHeader.classList.add('asc');
  };
  // ── State Restoration Functions ─────────────────────────────────────────
  const restoreFilterValue = () => {
    chrome.storage.local.get(
      ['profileFilterValue', 'currentDisplayType', 'targetMode', 'lastComplianceAssignments', 'lastAppAssignments', 'lastConfigAssignments', 'lastPwshAssignments', 'lastGroupMembers'],
      (data) => {
        // Restore target mode
        if (data.targetMode) {
          logMessage(`restoreFilterValue: Restoring target mode to '${data.targetMode}' from storage`);
          // Don't set state.targetMode first - let handleTargetModeToggle do it
          handleTargetModeToggle(data.targetMode);
        } else {
          logMessage(`restoreFilterValue: No stored target mode, using default '${state.targetMode}'`);
          // Ensure UI reflects the default state on first load
          handleTargetModeToggle(state.targetMode);
        }

        if (data.currentDisplayType) {
          state.currentDisplayType = data.currentDisplayType;
          clearTableAndPagination();
          if (state.currentDisplayType === 'compliance' && data.lastComplianceAssignments) {
            chrome.storage.local.remove(['lastConfigAssignments', 'lastAppAssignments', 'lastPwshAssignments']);
            updateComplianceTable(data.lastComplianceAssignments, false);
          } else if (state.currentDisplayType === 'apps' && data.lastAppAssignments) {
            chrome.storage.local.remove(['lastConfigAssignments', 'lastComplianceAssignments', 'lastPwshAssignments']);
            updateAppTable(data.lastAppAssignments, false);
          } else if (state.currentDisplayType === 'config' && data.lastConfigAssignments) {
            chrome.storage.local.remove(['lastAppAssignments', 'lastComplianceAssignments', 'lastPwshAssignments']);
            updateConfigTable(data.lastConfigAssignments, false);
          } else if (state.currentDisplayType === 'pwsh' && data.lastPwshAssignments) {
            chrome.storage.local.remove(['lastConfigAssignments', 'lastAppAssignments', 'lastComplianceAssignments', 'lastGroupMembers']);
            updatePwshTable(data.lastPwshAssignments, false);
          } else if (state.currentDisplayType === 'groupMembers' && data.lastGroupMembers) {
            chrome.storage.local.remove(['lastConfigAssignments', 'lastAppAssignments', 'lastComplianceAssignments', 'lastPwshAssignments']);
            updateGroupMembersTable(data.lastGroupMembers, false);
          }
        } else {
          clearTableAndPagination();
        }
        if (data.profileFilterValue) {
          document.getElementById('profileFilterInput').value = data.profileFilterValue;
          filterTable(data.profileFilterValue);
        }
      }
    );
  };
  const restoreState = () => {
    chrome.storage.local.get(['lastSearchResults', 'lastSearchQuery'], (data) => {
      if (data.lastSearchQuery) {
        document.getElementById("groupSearchInput").value = data.lastSearchQuery;
      }
      if (data.lastSearchResults) {
        const resultsDiv = document.getElementById("groupResults");
        resultsDiv.innerHTML = "";

        // Clear and rebuild dynamic groups tracking from stored results
        clearDynamicGroups();

        data.lastSearchResults.forEach(group => {
          // Restore dynamic group tracking
          if (group.isDynamic) {
            addDynamicGroup(group.id);
          }

          const item = document.createElement("p");
          const label = document.createElement("label");
          const checkbox = document.createElement("input");

          checkbox.type = "checkbox";
          checkbox.value = group.id;
          checkbox.id = "group-" + group.id;
          checkbox.className = "filled-in";
          checkbox.dataset.groupName = group.displayName;
          if (group.checked) checkbox.checked = true;

          if (group.isDynamic) {
            checkbox.title = "Dynamic group – cannot modify manually";
          }
          const span = document.createElement("span");
          span.textContent = group.displayName;
          if (group.isDynamic) {
            span.title = "Dynamic group – cannot modify manually";
          }


          label.appendChild(checkbox);
          label.appendChild(span);
          item.appendChild(label);
          resultsDiv.appendChild(item);
        });
        updateActionButtonsState();
      }
    });
  };
  // ── Event Handler Functions ─────────────────────────────────────────────
  // Handle Group Search
  const handleSearchGroup = async () => {
    logMessage("searchGroup clicked");
    const query = document.getElementById("groupSearchInput").value.trim();
    if (!query) {
      logMessage("searchGroup: No query entered");
      showNotification("Enter group name to search.", "info");
      return;
    }
    try {
      const token = await getToken();
      logMessage("searchGroup: Token found, proceeding with fetch");

      // Fetch groups with groupTypes information
      const groupsData = await fetchJSON(`https://graph.microsoft.com/v1.0/groups?$search="displayName:${query}"&$select=id,displayName,groupTypes&$top=10`, {
        method: "GET",
        headers: {
          "Authorization": token,
          "Content-Type": "application/json",
          "ConsistencyLevel": "eventual"
        }
      });

      logMessage("searchGroup: Received groups data");
      const resultsDiv = document.getElementById("groupResults");
      resultsDiv.innerHTML = "";

      if (!groupsData.value || groupsData.value.length === 0) {
        logMessage("searchGroup: No groups found");
        resultsDiv.textContent = "No groups found.";
        return;
      }

      // Clear previous dynamic groups tracking for search results
      clearDynamicGroups();

      const searchResults = groupsData.value.map(g => {
        const isDynamic = g.groupTypes && g.groupTypes.includes('DynamicMembership');
        if (isDynamic) {
          addDynamicGroup(g.id);
        }

        return {
          id: g.id,
          displayName: g.displayName,
          isDynamic: isDynamic,
          checked: false
        };
      });

      searchResults.forEach(group => {
        const item = document.createElement("p");
        const label = document.createElement("label");
        const checkbox = document.createElement("input");

        checkbox.type = "checkbox";
        checkbox.value = group.id;
        checkbox.id = "group-" + group.id;
        checkbox.className = "filled-in";
        checkbox.dataset.groupName = group.displayName;

        if (group.isDynamic) {
          checkbox.title = "Dynamic group – cannot modify manually";
        }

        const span = document.createElement("span");
        span.textContent = group.displayName;
        if (group.isDynamic) {
          span.title = "Dynamic group – cannot modify manually";
        }

        label.appendChild(checkbox);
        label.appendChild(span);
        item.appendChild(label); resultsDiv.appendChild(item);
      });
      updateActionButtonsState();

      chrome.storage.local.set({ lastSearchResults: searchResults, lastSearchQuery: query });
    } catch (error) {
      logMessage(`searchGroup: Error - ${error.message}`);
      showNotification('Error: ' + error.message, 'error');
    }
  };

  // Handle Adding Device/User to Selected Groups
  const handleAddToGroups = async () => {
    const targetType = state.targetMode === 'device' ? 'device' : 'user';
    logMessage(`addToGroups clicked (${targetType} mode)`);

    if (document.getElementById('addToGroups').classList.contains('disabled')) {
      showNotification('Cannot modify dynamic groups.', 'error');
      return;
    }

    const allSelected = getAllSelectedGroups();
    if (!allSelected.hasAnySelection) {
      logMessage("addToGroups: No groups selected");
      alert("Select at least one group from search results or table rows.");
      return;
    }

    // Show notification with count
    const totalCount = allSelected.searchResults.length + allSelected.tableSelections.length;
    showNotification(`Adding ${targetType} to ${totalCount} group(s)...`, 'info');

    try {
      const { mdmDeviceId } = await verifyMdmUrl();
      const token = await getToken();
      logMessage(`addToGroups: Extracted mdmDeviceId ${mdmDeviceId}`);

      const { directoryId, displayName } = await getDirectoryObjectId(mdmDeviceId, token, state.targetMode);
      logMessage(`addToGroups: Got directory ID for ${targetType}: ${directoryId}`);

      const promises = [];

      // Process search result groups (existing logic with group IDs)
      allSelected.searchResults.forEach(group => {
        promises.push(addToSingleGroup(group.id, directoryId, token, group.name));
      });

      // Process table selection groups (need to resolve names to IDs)
      for (const groupName of allSelected.tableSelections) {
        try {
          const groupData = await fetchJSON(`https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(groupName)}'&$select=id,displayName`, {
            method: "GET",
            headers: { "Authorization": token, "Content-Type": "application/json" }
          });

          if (groupData.value && groupData.value.length > 0) {
            const groupId = groupData.value[0].id;
            promises.push(addToSingleGroup(groupId, directoryId, token, groupName));
          } else {
            promises.push(Promise.resolve({ groupName, error: "Group not found" }));
          }
        } catch (e) {
          promises.push(Promise.resolve({ groupName, error: e.message }));
        }
      }

      const results = await Promise.all(promises);
      let success = true;
      let message = '';
      results.forEach(r => {
        if (r.error) {
          success = false;
          const identifier = r.groupId || r.groupName || 'Unknown Group';
          message += `Group ${identifier}: Error - ${r.error}\n`;
        } else {
          const identifier = r.groupId || r.groupName || 'Unknown Group';
          message += `Group ${identifier}: Success\n`;
        }
      });

      const successMsg = `Successfully added ${targetType} "${displayName}" to groups`;
      const errorMsg = `Some groups could not be added for ${targetType} "${displayName}"\n${message}`;
      showNotification(success ? successMsg : errorMsg, success ? 'success' : 'error');
    } catch (error) {
      logMessage(`addToGroups: Error - ${error.message}`);
      showNotification(`Failed to add ${targetType} to groups: ${error.message}`, 'error');
    }
  };

  // Helper function to add to a single group
  const addToSingleGroup = async (groupId, directoryId, token, groupName = null) => {
    const postUrl = `https://graph.microsoft.com/beta/groups/${groupId}/members/$ref`;
    const body = JSON.stringify({
      "@odata.id": `https://graph.microsoft.com/beta/directoryObjects/${directoryId}`
    });

    try {
      await fetchJSON(postUrl, {
        method: "POST",
        headers: { "Authorization": token, "Content-Type": "application/json" },
        body: body
      });
      return { groupId, groupName, result: "Success" };
    } catch (e) {
      return { groupId, groupName, error: e.message };
    }
  };  // Handle Removing Device/User from Selected Groups
  const handleRemoveFromGroups = async () => {
    const targetType = state.targetMode === 'device' ? 'device' : 'user';
    logMessage(`removeFromGroups clicked (${targetType} mode)`);

    if (document.getElementById('removeFromGroups').classList.contains('disabled')) {
      showNotification('Cannot modify dynamic groups.', 'error');
      return;
    }

    const allSelected = getAllSelectedGroups();
    if (!allSelected.hasAnySelection) {
      logMessage("removeFromGroups: No groups selected");
      alert("Select at least one group from search results or table rows.");
      return;
    }

    // Show notification with count
    const totalCount = allSelected.searchResults.length + allSelected.tableSelections.length;
    showNotification(`Removing ${targetType} from ${totalCount} group(s)...`, 'info');

    try {
      const { mdmDeviceId } = await verifyMdmUrl();
      const token = await getToken();
      logMessage(`removeFromGroups: Extracted mdmDeviceId ${mdmDeviceId}`);

      const { directoryId, displayName } = await getDirectoryObjectId(mdmDeviceId, token, state.targetMode);
      logMessage(`removeFromGroups: Got directory ID for ${targetType}: ${directoryId}`);

      const promises = [];

      // Process search result groups (existing logic with group IDs)
      allSelected.searchResults.forEach(group => {
        promises.push(removeFromSingleGroup(group.id, directoryId, token, group.name));
      });

      // Process table selection groups (need to resolve names to IDs)
      for (const groupName of allSelected.tableSelections) {
        try {
          const groupData = await fetchJSON(`https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(groupName)}'&$select=id,displayName`, {
            method: "GET",
            headers: { "Authorization": token, "Content-Type": "application/json" }
          });

          if (groupData.value && groupData.value.length > 0) {
            const groupId = groupData.value[0].id;
            promises.push(removeFromSingleGroup(groupId, directoryId, token, groupName));
          } else {
            promises.push(Promise.resolve({ groupName, error: "Group not found" }));
          }
        } catch (e) {
          promises.push(Promise.resolve({ groupName, error: e.message }));
        }
      }

      const results = await Promise.all(promises);
      let success = true;
      let message = '';
      results.forEach(r => {
        if (r.error) {
          success = false;
          const identifier = r.groupId || r.groupName || 'Unknown Group';
          message += `Group ${identifier}: Error - ${r.error}\n`;
        } else {
          const identifier = r.groupId || r.groupName || 'Unknown Group';
          message += `Group ${identifier}: Success\n`;
        }
      });

      const successMsg = `Successfully removed ${targetType} "${displayName}" from groups`;
      const errorMsg = `Some groups could not be removed for ${targetType} "${displayName}"\n${message}`;
      showNotification(success ? successMsg : errorMsg, success ? 'success' : 'error');
    } catch (error) {
      logMessage(`removeFromGroups: Error - ${error.message}`);
      showNotification(`Failed to remove ${targetType} from groups: ${error.message}`, 'error');
    }
  };

  // Helper function to remove from a single group
  const removeFromSingleGroup = async (groupId, directoryId, token, groupName = null) => {
    const deleteUrl = `https://graph.microsoft.com/v1.0/groups/${groupId}/members/${directoryId}/$ref`;

    try {
      const response = await fetch(deleteUrl, {
        method: "DELETE",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      if (response.ok) return { groupId, groupName, result: "Removed" };
      const text = await response.text();
      return { groupId, groupName, error: text || response.statusText };
    } catch (e) {
      return { groupId, groupName, error: e.message };
    }
  };

  const fetchAllGroupMembers = async (groupId, token) => {
    let members = [];
    let url = `https://graph.microsoft.com/v1.0/groups/${groupId}/transitiveMembers?$select=id,displayName,userPrincipalName,deviceId,@odata.type`;
    while (url) {
      const data = await fetchJSON(url, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      if (data.value) members = members.concat(data.value);
      url = data['@odata.nextLink'] || null;
    }
    return members.filter(m => m['@odata.type'] !== '#microsoft.graph.group');
  };

  // Handle Checking Group Members
  const handleCheckGroupMembers = async () => {
    logMessage("checkGroupMembers clicked");
    const selected = document.querySelectorAll("#groupResults input[type=checkbox]:checked");
    if (selected.length !== 1) {
      showNotification('Select exactly one group.', 'error');
      return;
    }

    document.getElementById('profileFilterInput').value = '';
    chrome.storage.local.set({ profileFilterValue: '' });

    clearTableSelection();

    const groupId = selected[0].value;
    const groupName = selected[0].dataset.groupName;

    try {
      const token = await getToken();
      const members = await fetchAllGroupMembers(groupId, token);

      chrome.storage.local.remove(['lastConfigAssignments','lastAppAssignments','lastComplianceAssignments','lastPwshAssignments']);
      chrome.storage.local.set({ lastGroupMembers: members });

      document.getElementById('deviceNameDisplay').textContent = `- ${groupName} (${members.length} members)`;

      updateGroupMembersTable(members);

      showNotification('Group members loaded successfully', 'success');
    } catch (error) {
      logMessage(`checkGroupMembers: Error - ${error.message}`);
      showNotification('Failed to load group members: ' + error.message, 'error');
    }
  };

  // Handle Checking Configuration Assignments
  const handleCheckGroups = async () => {
    logMessage("checkGroups clicked");
    showNotification('Fetching configuration assignments...', 'info');
    document.getElementById('profileFilterInput').value = '';
    chrome.storage.local.set({ profileFilterValue: '' });

    chrome.storage.local.remove(['lastAppAssignments', 'lastComplianceAssignments', 'lastPwshAssignments', 'lastGroupMembers']);

    // Clear table selection before loading new assignments
    clearTableSelection();

    // Clear dynamic groups tracking before fetching new assignments
    clearDynamicGroups();

    try {
      const { mdmDeviceId } = await verifyMdmUrl();
      const token = await getToken();
      const deviceData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/manageddevices('${mdmDeviceId}')`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      // Update device name display
      updateDeviceNameDisplay(deviceData);

      const azureADDeviceId = deviceData.azureADDeviceId;
      const userPrincipalName = deviceData.userPrincipalName;
      if (!azureADDeviceId) throw new Error("Could not find Azure AD Device ID for this device.");
      const filter = encodeURIComponent(`deviceId eq '${azureADDeviceId}'`);
      const deviceObjData = await fetchJSON(`https://graph.microsoft.com/beta/devices?$filter=${filter}`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      if (!deviceObjData.value || deviceObjData.value.length === 0) throw new Error("Could not find the device in Azure AD.");
      const deviceObjectId = deviceObjData.value[0].id;
      let userPromise;
      if (userPrincipalName && userPrincipalName !== 'Unknown user') {
        userPromise = fetchJSON(`https://graph.microsoft.com/beta/users?$filter=userPrincipalName eq '${encodeURIComponent(userPrincipalName)}'`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        }).then(userData => (userData.value && userData.value.length > 0) ? userData.value[0].id : null);
      } else {
        userPromise = Promise.resolve(null);
      }
      const userObjectId = await userPromise;
      const allGroups = await getAllGroupsMap(deviceObjectId, userObjectId, token);
      const reportBody = JSON.stringify({
        top: "500",
        skip: "0",
        select: ["PolicyId", "PolicyName", "PolicyType", "UPN"],
        filter: `((PolicyBaseTypeName eq 'Microsoft.Management.Services.Api.DeviceConfiguration') or (PolicyBaseTypeName eq 'DeviceManagementConfigurationPolicy') or (PolicyBaseTypeName eq 'DeviceConfigurationAdmxPolicy') or (PolicyBaseTypeName eq 'Microsoft.Management.Services.Api.DeviceManagementIntent')) and (IntuneDeviceId eq '${mdmDeviceId}')`
      });
      const reportData = await fetchJSON("https://graph.microsoft.com/beta/deviceManagement/reports/getConfigurationPoliciesReportForDevice", {
        method: "POST",
        headers: { "Authorization": token, "Content-Type": "application/json" },
        body: reportBody
      });
      const policies = reportData.Values || [];
      if (policies.length === 0) {
        showNotification('No policies found.', 'info');
        clearTableAndPagination();
        return;
      }
      const assignmentsPromises = policies.map(async (row) => {
        const policy = {
          PolicyId: row[0],
          PolicyName: row[1],
          PolicyType: row[2],
          UPN: row[3]
        };
        const specialPolicyIds = ["26", "20", "33", "55", "118", "75", "72", "25", "31", "107", "99999"];
        const endpoint = specialPolicyIds.includes(policy.PolicyType.toString())
          ? `https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/${policy.PolicyId}/assignments`
          : `https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('${policy.PolicyId}')/assignments`;
        try {
          const assignData = await fetchJSON(endpoint, {
            method: "GET",
            headers: { "Authorization": token, "Content-Type": "application/json" }
          });
          let targetObjs = [];
          const assignments = Array.isArray(assignData.value) ? assignData.value : (assignData.value ? [assignData.value] : []);
          assignments.forEach(asg => {
            if (!asg.target) return;
            const typeRaw = (asg.target['@odata.type'] || "").toLowerCase().trim();
            if (typeRaw.includes("groupassignmenttarget")) {
              if (allGroups.has(asg.target.groupId)) {
                targetObjs.push({
                  groupId: asg.target.groupId,
                  groupName: allGroups.get(asg.target.groupId),
                  membershipType: 'Direct',
                  targetType: typeRaw.includes('user') ? 'User' : 'Device',
                  intent: "Included"
                });
              }
            } else if (typeRaw.includes("alldevicesassignmenttarget")) {
              targetObjs.push({
                groupName: "All Devices",
                membershipType: "Virtual",
                targetType: "Device",
                intent: "Included"
              });
            } else if (typeRaw.includes("allusersassignmenttarget")) {
              if (policy.UPN && policy.UPN !== 'Not Available') {
                targetObjs.push({
                  groupName: "All Users",
                  membershipType: "Virtual",
                  targetType: "User",
                  intent: "Included"
                });
              }
            }
          });
          return { policyName: policy.PolicyName, targets: targetObjs };
        } catch (err) {
          logMessage(`checkGroups: Error processing policy ${policy.PolicyName} - ${err.message}`);
          return { policyName: policy.PolicyName, targets: [] };
        }
      });
      const finalResults = (await Promise.all(assignmentsPromises)).filter(result => result.targets && result.targets.length > 0);
      chrome.storage.local.set({ lastConfigAssignments: finalResults });
      updateConfigTable(finalResults);
      logMessage(`checkGroups: Found ${finalResults.length} policies with valid group assignments`);
      showNotification('Configuration assignments loaded successfully', 'success');
    } catch (error) {
      logMessage(`checkGroups: Error - ${error.message}`);
      showNotification('Failed to load configuration assignments: ' + error.message, 'error');
    }
  };
  // Handle Checking Compliance Policies
  const handleCheckCompliance = async () => {
    logMessage("checkCompliance clicked");
    showNotification('Fetching compliance policies...', 'info');
    document.getElementById('profileFilterInput').value = '';
    chrome.storage.local.set({ profileFilterValue: '' });

    chrome.storage.local.remove(['lastConfigAssignments','lastAppAssignments','lastPwshAssignments','lastGroupMembers']);

    // Clear table selection before loading new assignments
    clearTableSelection();
    let token, mdmDeviceId;
    try {
      ({ mdmDeviceId } = await verifyMdmUrl());
      token = await getToken();
      const deviceData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/manageddevices('${mdmDeviceId}')`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      // Update device name display
      updateDeviceNameDisplay(deviceData);

      const azureADDeviceId = deviceData.azureADDeviceId;
      const userPrincipalName = deviceData.userPrincipalName;
      if (!azureADDeviceId) throw new Error("Could not find Azure AD Device ID for this device.");
      const filter = encodeURIComponent(`deviceId eq '${azureADDeviceId}'`);
      const deviceObjData = await fetchJSON(`https://graph.microsoft.com/beta/devices?$filter=${filter}`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      if (!deviceObjData.value || deviceObjData.value.length === 0) throw new Error("Could not find the device in Azure AD.");
      const deviceObjectId = deviceObjData.value[0].id;
      let userPromise;
      if (userPrincipalName && userPrincipalName !== 'Unknown user') {
        userPromise = fetchJSON(`https://graph.microsoft.com/beta/users?$filter=userPrincipalName eq '${encodeURIComponent(userPrincipalName)}'`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        }).then(userData => (userData.value && userData.value.length > 0) ? userData.value[0].id : null);
      } else {
        userPromise = Promise.resolve(null);
      }
      const userObjectId = await userPromise;
      const allGroups = await getAllGroupsMap(deviceObjectId, userObjectId, token);
      const reportBody = JSON.stringify({
        filter: `(DeviceId eq '${mdmDeviceId}') and ((PolicyPlatformType eq '4') or (PolicyPlatformType eq '5') or (PolicyPlatformType eq '6') or (PolicyPlatformType eq '8') or (PolicyPlatformType eq '100'))`,
        orderBy: ["PolicyName asc"]
      });
      const reportData = await fetchJSON("https://graph.microsoft.com/beta/deviceManagement/reports/getDevicePoliciesComplianceReport", {
        method: "POST",
        headers: { "Authorization": token, "Content-Type": "application/json" },
        body: reportBody
      });
      if (!reportData || !reportData.Schema || !reportData.Values) {
        logMessage("checkCompliance: Invalid response format");
        showNotification('Invalid response format from compliance report API.', 'error');
        return;
      }
      const schemaMap = {};
      reportData.Schema.forEach((col, idx) => { schemaMap[col.Column] = idx; });
      const requiredColumns = ['PolicyId', 'PolicyName', 'PolicyStatus_loc'];
      const missingColumns = requiredColumns.filter(col => schemaMap[col] === undefined);
      if (missingColumns.length > 0) {
        logMessage(`checkCompliance: Missing required columns: ${missingColumns.join(', ')}`);
        showNotification(`API response missing required columns: ${missingColumns.join(', ')}`, 'error');
        return;
      }
      const policies = reportData.Values || [];
      if (policies.length === 0) {
        logMessage("checkCompliance: No compliance policies found");
        showNotification('No compliance policies found.', 'info');
        clearTableAndPagination();
        return;
      }
      const formattedPolicies = policies.map(policy => ({
        policyId: policy[schemaMap.PolicyId] || 'Unknown',
        policyName: policy[schemaMap.PolicyName] || 'Unknown Policy',
        complianceState: policy[schemaMap.PolicyStatus_loc] || 'Unknown'
      }));
      const assignmentPromises = formattedPolicies.map(async (policy) => {
        if (!policy.policyId || policy.policyId === 'Unknown') {
          return { ...policy, assignments: [] };
        }
        try {
          const policyData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/${policy.policyId}?$expand=assignments`, {
            method: "GET",
            headers: { "Authorization": token, "Content-Type": "application/json" }
          });
          return { ...policy, assignments: policyData.assignments || [] };
        } catch (err) {
          logMessage(`checkCompliance: Error getting assignments for policy ${policy.policyId}: ${err.message}`);
          return { ...policy, assignments: [] };
        }
      });
      const policiesWithAssignments = await Promise.all(assignmentPromises);
      if (!policiesWithAssignments || policiesWithAssignments.length === 0) {
        clearTableAndPagination();
        showNotification('No compliance policies found.', 'info');
        return;
      }
      const tableData = policiesWithAssignments.map(policy => {
        let targets = [];
        if (policy.assignments && policy.assignments.length > 0) {
          policy.assignments.forEach(asg => {
            if (!asg.target) return;
            const targetType = (asg.target['@odata.type'] || '').toLowerCase();
            const isExclusion = targetType.includes('exclusion');
            if (isExclusion) return; // Skip exclusions

            if (targetType.includes('groupassignmenttarget')) {
              const groupId = asg.target.groupId;
              const groupName = allGroups.has(groupId)
                ? allGroups.get(groupId)
                : `Group ID: ${groupId.substring(0, 8)}...`;

              // Skip unresolved group IDs
              if (groupName.startsWith('Group ID:')) return;

              targets.push({
                groupName,
                membershipType: isExclusion ? 'Exclude' : 'Direct',
                targetType: targetType.includes('user') ? 'User' : 'Device'
              });
            } else if (targetType.includes('alldevicesassignmenttarget')) {
              targets.push({
                groupName: 'All Devices',
                membershipType: isExclusion ? 'Exclude' : 'Virtual',
                targetType: 'Device'
              });
            } else if (targetType.includes('allusersassignmenttarget')) {
              targets.push({
                groupName: 'All Users',
                membershipType: isExclusion ? 'Exclude' : 'Virtual',
                targetType: 'User'
              });
            }
          });
        }
        if (targets.length === 0) {
          targets.push({ groupName: 'No Assignments', membershipType: '-', targetType: '-' });
        }
        return { policyName: policy.policyName, complianceState: policy.complianceState, targets };
      });
      chrome.storage.local.set({ lastComplianceAssignments: tableData });
      updateComplianceTable(tableData);
      logMessage(`checkCompliance: Found ${tableData.length} compliance policies with assignments`);
      showNotification('Compliance policies loaded successfully', 'success');
    } catch (error) {
      logMessage(`checkCompliance: Error - ${error.message}`);
      showNotification('Failed to load compliance policies: ' + error.message, 'error');
    }
  };

  // Handle Downloading Script
  const handleDownloadScript = async () => {
    logMessage("downloadScript clicked");
    chrome.tabs.query({ active: true, currentWindow: true }, async (tabs) => {
      if (!tabs || !tabs[0]) {
        logMessage("downloadScript: No active tab found");
        showNotification('No active tab found.', 'error');
        return;
      }
      const url = tabs[0].url;
      const policyMatch = url.match(/policyId\/([\w-]+)/);
      if (!policyMatch || !policyMatch[1]) {
        logMessage("downloadScript: No policyId found in URL");
        showNotification('Could not find policy ID in the current URL.', 'error');
        return;
      }
      const policyId = policyMatch[1];
      logMessage(`downloadScript: Found policyId: ${policyId}`);
      try {
        const token = await getToken();
        const scriptData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts/${policyId}`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        });
        if (!scriptData.scriptContent) {
          throw new Error("No script content found in response");
        }
        const decodedScript = atob(scriptData.scriptContent);
        const blob = new Blob([decodedScript], { type: 'text/plain' });
        const blobUrl = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = blobUrl;
        a.download = scriptData.fileName || 'IntuneScript.ps1';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(blobUrl);
        document.body.removeChild(a);
        showNotification('Script downloaded successfully', 'success');
      } catch (error) {
        logMessage(`downloadScript: Error - ${error.message}`);
        showNotification('Failed to download script: ' + error.message, 'error');
      }
    });
  };
  // Handle App Assignments
  const handleAppsAssignment = async () => {
    logMessage("appsAssignment clicked");
    showNotification('Fetching app assignments...', 'info');
    document.getElementById('profileFilterInput').value = '';
    chrome.storage.local.set({ profileFilterValue: '' });

    chrome.storage.local.remove(['lastConfigAssignments','lastComplianceAssignments','lastPwshAssignments','lastGroupMembers']);

    // Clear table selection before loading new assignments
    clearTableSelection();
    let token;
    try {
      const { mdmDeviceId } = await verifyMdmUrl();
      token = await getToken();
      const deviceData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/manageddevices('${mdmDeviceId}')`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      // Update device name display
      updateDeviceNameDisplay(deviceData);

      const azureADDeviceId = deviceData.azureADDeviceId;
      const userPrincipalName = deviceData.userPrincipalName;
      if (!azureADDeviceId) throw new Error("Could not find Azure AD Device ID for this device.");
      const filter = encodeURIComponent(`deviceId eq '${azureADDeviceId}'`);
      const deviceObjData = await fetchJSON(`https://graph.microsoft.com/beta/devices?$filter=${filter}`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      if (!deviceObjData.value || deviceObjData.value.length === 0) throw new Error("Could not find the device in Azure AD.");
      const deviceObjectId = deviceObjData.value[0].id;
      let userObjectId = null;
      if (userPrincipalName && userPrincipalName !== 'Unknown user') {
        const userData = await fetchJSON(`https://graph.microsoft.com/beta/users?$filter=userPrincipalName eq '${encodeURIComponent(userPrincipalName)}'`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        });
        if (userData.value && userData.value.length > 0) {
          userObjectId = userData.value[0].id;
        }
      }
      const allGroups = await getAllGroupsMap(deviceObjectId, userObjectId, token);
      const appRequests = [];
      appRequests.push(
        fetchJSON(`https://graph.microsoft.com/beta/users('00000000-0000-0000-0000-000000000000')/mobileAppIntentAndStates('${mdmDeviceId}')`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        }).then(data => ({ type: 'Device', apps: data.mobileAppList || [] }))
      );
      if (userObjectId) {
        appRequests.push(
          fetchJSON(`https://graph.microsoft.com/beta/users('${userObjectId}')/mobileAppIntentAndStates('${mdmDeviceId}')`, {
            method: "GET",
            headers: { "Authorization": token, "Content-Type": "application/json" }
          }).then(data => ({ type: 'User', apps: data.mobileAppList || [] }))
        );
      }
      const appResults = await Promise.all(appRequests);
      const allApps = [];
      let totalApps = 0;
      appResults.forEach(result => {
        totalApps += result.apps.length;
        result.apps.forEach(app => {
          allApps.push({
            applicationId: app.applicationId,
            displayName: app.displayName,
            mobileAppIntent: app.mobileAppIntent,
            displayVersion: app.displayVersion || 'N/A',
            installState: app.installState,
            targetType: result.type
          });
        });
      });
      logMessage(`appsAssignment: Found ${totalApps} total apps`);
      const assignmentPromises = allApps.map(async (app) => {
        const assignmentData = await fetchJSON(`https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/${app.applicationId}/assignments`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        }).catch(() => ({ value: [] }));
        const assignments = assignmentData.value || [];
        const validTargets = [];
        assignments.forEach(assignment => {
          if (!assignment.target) return;
          const intentInfo = assignment.intent || app.mobileAppIntent;
          const typeRaw = (assignment.target['@odata.type'] || '').toLowerCase();

          // Skip exclusions
          if (typeRaw.includes('exclusion')) return;

          if (typeRaw.includes('alldevicesassignmenttarget')) {
            validTargets.push({
              groupName: 'All Devices',
              membershipType: 'Virtual',
              targetType: 'Device',
              intent: intentInfo
            });
          } else if (typeRaw.includes('allusersassignmenttarget') || typeRaw.includes('alllicensedusersassignmenttarget')) {
            validTargets.push({
              groupName: 'All Users',
              membershipType: 'Virtual',
              targetType: 'User',
              intent: intentInfo
            });
          } else if (typeRaw.includes('groupassignmenttarget')) {
            const groupId = assignment.target.groupId;
            const groupName = allGroups.has(groupId) ? allGroups.get(groupId) : 'Group ID: ' + groupId.substring(0, 8) + '...';

            // Skip unresolved group IDs
            if (groupName.startsWith('Group ID:')) return;

            validTargets.push({
              groupId,
              groupName,
              membershipType: 'Direct',
              targetType: typeRaw.includes('user') ? 'User' : 'Device',
              intent: intentInfo
            });
          } else {
            validTargets.push({
              groupName: typeRaw,
              membershipType: '-',
              targetType: '-',
              intent: intentInfo
            });
          }
        });
        if (validTargets.length === 0) {
          validTargets.push({
            groupName: 'No Assignments',
            membershipType: '-',
            targetType: app.targetType,
            intent: app.mobileAppIntent || '-'
          });
        }
        return {
          appName: app.displayName,
          appVersion: app.displayVersion,
          installState: app.installState,
          targets: validTargets
        };
      });
      const appAssignments = await Promise.all(assignmentPromises);
      chrome.storage.local.set({ lastAppAssignments: appAssignments });
      updateAppTable(appAssignments);
      logMessage(`appsAssignment: Found ${appAssignments.length} apps total`);
      showNotification('App assignments loaded successfully', 'success');
    } catch (error) {
      logMessage(`appsAssignment: Error - ${error.message}`);
      showNotification('Failed to load app assignments: ' + error.message, 'error');
    }
  };
  // Handle PowerShell Profiles (scripts)
  const handlePwshProfiles = async () => {
    logMessage("pwshProfiles clicked");
    showNotification('Fetching PowerShell profiles...', 'info');
    document.getElementById('profileFilterInput').value = '';
    chrome.storage.local.set({ profileFilterValue: '' });

    chrome.storage.local.remove(['lastConfigAssignments','lastAppAssignments','lastComplianceAssignments','lastGroupMembers']);

    // Clear table selection before loading new assignments
    clearTableSelection();
    let token;
    try {
      const { mdmDeviceId } = await verifyMdmUrl();
      token = await getToken();
      const deviceData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/manageddevices('${mdmDeviceId}')`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      // Update device name display
      updateDeviceNameDisplay(deviceData);

      const azureADDeviceId = deviceData.azureADDeviceId;
      const userPrincipalName = deviceData.userPrincipalName;
      if (!azureADDeviceId) throw new Error("Could not find Azure AD Device ID for this device.");
      const filter = encodeURIComponent(`deviceId eq '${azureADDeviceId}'`);
      const deviceObjData = await fetchJSON(`https://graph.microsoft.com/beta/devices?$filter=${filter}`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      if (!deviceObjData.value || deviceObjData.value.length === 0) throw new Error("Could not find the device in Azure AD.");
      const deviceObjectId = deviceObjData.value[0].id;
      let userObjectId = null;
      if (userPrincipalName && userPrincipalName !== 'Unknown user') {
        const userData = await fetchJSON(`https://graph.microsoft.com/beta/users?$filter=userPrincipalName eq '${encodeURIComponent(userPrincipalName)}'`, {
          method: "GET",
          headers: { "Authorization": token, "Content-Type": "application/json" }
        });
        if (userData.value && userData.value.length > 0) userObjectId = userData.value[0].id;
      }
      const allGroups = await getAllGroupsMap(deviceObjectId, userObjectId, token);
      const scriptsData = await fetchJSON("https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts?$expand=assignments", {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });
      if (!scriptsData.value || scriptsData.value.length === 0) {
        showNotification('No PowerShell profiles found.', 'info');
        clearTableAndPagination();
        return;
      }
      const matchedScripts = [];
      let matchCount = 0;
      scriptsData.value.forEach(script => {
        if (!script.assignments || script.assignments.length === 0) {
          logMessage(`pwshProfiles: Script "${script.displayName}" has no assignments - skipping`);
          return;
        }
        script.assignments.forEach(asg => {
          if (!asg.target) {
            logMessage(`pwshProfiles: Script "${script.displayName}" has assignment without target - skipping`);
            return;
          }
          let targetName = '';
          let isMatch = false;
          if (asg.target.groupId) {
            if (allGroups.has(asg.target.groupId)) {
              targetName = allGroups.get(asg.target.groupId);
              isMatch = true;
              matchCount++;
              logMessage(`pwshProfiles: MATCH - Group ${targetName}`);
            } else {
              logMessage(`pwshProfiles: NO MATCH - Group ID ${asg.target.groupId} not in user/device groups - skipping`);
              return;
            }
          } else if (asg.target['@odata.type']) {
            const targetType = asg.target['@odata.type'].toLowerCase();
            if (targetType.includes('alldevicesassignmenttarget')) {
              targetName = 'All Devices';
              isMatch = true;
              matchCount++;
              logMessage(`pwshProfiles: MATCH - All Devices`);
            } else if (targetType.includes('allusersassignmenttarget') || targetType.includes('alllicensedusersassignmenttarget')) {
              targetName = 'All Users';
              isMatch = true;
              matchCount++;
              logMessage(`pwshProfiles: MATCH - All Users`);
            } else {
              logMessage(`pwshProfiles: UNKNOWN target type ${targetType} - skipping`);
              return;
            }
          } else {
            logMessage(`pwshProfiles: Assignment has no target info - skipping`);
            return;
          }
          if (isMatch) {
            matchedScripts.push({
              scriptName: script.displayName,
              description: script.description || '',
              targetName,
              targetGroupId: asg.target.groupId || null // Include group ID for dynamic group checking
            });
          }
        });
      });
      chrome.storage.local.set({ lastPwshAssignments: matchedScripts });
      updatePwshTable(matchedScripts);
      logMessage(`pwshProfiles: Found ${matchCount} matching assignments, saved ${matchedScripts.length} script entries`);
      if (matchCount === 0) {
        showNotification('No matching PowerShell profiles found for this device/user.', 'info');
      } else {
        showNotification(`PowerShell profiles loaded. Found ${matchCount} matches.`, 'success');
      }
    } catch (error) {
      logMessage(`pwshProfiles: Error - ${error.message}`);
      showNotification('Failed to load PowerShell profiles: ' + error.message, 'error');
    }
  };

  // Handle Group Creation
  const handleCreateGroup = async () => {
    logMessage("createGroup clicked");
    const groupName = document.getElementById("groupSearchInput").value.trim();
    if (!groupName) {
      showNotification('Please enter a group name.', 'error');
      return;
    }
    const mailNickname = groupName.substring(0, 10).replace(/[^a-zA-Z0-9]/g, '');
    chrome.storage.local.get("msGraphToken", async (data) => {
      if (!data.msGraphToken) {
        showNotification('No token captured. Please login first.', 'error');
        return;
      }
      try {
        const result = await fetchJSON('https://graph.microsoft.com/beta/groups', {
          method: 'POST',
          headers: {
            'Authorization': data.msGraphToken,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            displayName: groupName,
            mailEnabled: false,
            mailNickname: mailNickname,
            securityEnabled: true
          })
        });
        logMessage("Group created successfully");
        showNotification('Group created successfully!', 'info');
        // Refresh group list
        document.getElementById("searchGroup").click();
      } catch (error) {
        logMessage(`createGroup: Error - ${error.message}`);
        showNotification('Error creating group: ' + error.message, 'error');
      }
    });
  };

  // Handle Collecting Log Files
  const handleCollectLogs = async () => {
    logMessage("collectLogs clicked");

    // Prompt user for log paths and AppID
    const logPathsPrompt = "Enter log paths (supported folders: %PROGRAMFILES%, %PROGRAMDATA%, %PUBLIC%, %WINDIR%, %TEMP%, %TMP%). Delimit paths with semicolons (;)";

    const logPaths = prompt(logPathsPrompt, "%PROGRAMDATA%\\Microsoft\\IntuneManagementExtension\\Logs\\IntuneManagementExtension.log");

    if (!logPaths) {
      logMessage("collectLogs: No log paths entered");
      showNotification("Log collection canceled.", "info");
      return;
    }

    const appIdPrompt = "Enter Intune Win32 application ID to initialize log collection (recommend using a dummy app ID)";
    const appId = prompt(appIdPrompt, "");

    if (!appId) {
      logMessage("collectLogs: No AppID entered");
      showNotification("AppID is required for log collection.", "error");
      return;
    }

    try {
      const { mdmDeviceId } = await verifyMdmUrl();
      const token = await getToken();
      // Get device data to find the primary user
      const deviceData = await fetchJSON(`https://graph.microsoft.com/beta/deviceManagement/manageddevices('${mdmDeviceId}')`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      // Update device name display
      updateDeviceNameDisplay(deviceData);

      const userPrincipalName = deviceData.userPrincipalName;
      if (!userPrincipalName || userPrincipalName === 'Unknown user') {
        throw new Error("No primary user found for this device.");
      }

      // Get the user ID
      const userData = await fetchJSON(`https://graph.microsoft.com/beta/users?$filter=userPrincipalName eq '${encodeURIComponent(userPrincipalName)}'`, {
        method: "GET",
        headers: { "Authorization": token, "Content-Type": "application/json" }
      });

      if (!userData.value || userData.value.length === 0) {
        throw new Error("User not found in Azure AD.");
      }

      const userObjectId = userData.value[0].id;

      // Format the log paths for the API - with proper escaping
      const logPathsArray = [];
      logPaths.split(';').forEach(path => {
        const trimmedPath = path.trim();
        if (trimmedPath.length > 0) {
          // Replace single backslashes with double backslashes
          // But don't stringify yet to avoid extra escaping
          const escapedPath = trimmedPath.replace(/\\/g, '\\\\');
          logPathsArray.push(`"${escapedPath}"`);
        }
      });

      if (logPathsArray.length === 0) {
        throw new Error("No valid log paths provided.");
      }

      // Create raw JSON string with exactly the format needed
      // Note: We're manually constructing the JSON to avoid extra escaping by JSON.stringify
      const rawJsonBody = `{
        "customLogFolders": [${logPathsArray.join(', ')}],
        "id": "${userObjectId}_${mdmDeviceId}_${appId}"
      }`;

      logMessage(`collectLogs: Log paths: ${JSON.stringify(logPathsArray)}`);
      logMessage(`collectLogs: Requesting logs for user ${userObjectId}, device ${mdmDeviceId}, app ${appId}`);

      // Make the API call
      const requestUrl = `https://graph.microsoft.com/beta/users('${userObjectId}')/mobileAppTroubleshootingEvents('${mdmDeviceId}_${appId}')/appLogCollectionRequests`;

      const logResult = await fetchJSON(requestUrl, {
        method: "POST",
        headers: {
          "Authorization": token,
          "Content-Type": "application/json"
        },
        body: rawJsonBody
      });

      logMessage("collectLogs: Log collection request successful");
      showNotification("Log collection initiated successfully. The logs will be collected on the device and uploaded to Intune.", "success");
    } catch (error) {
      logMessage(`collectLogs: Error - ${error.message}`);
      showNotification('Failed to collect logs: ' + error.message, 'error');
    }
  };

  // ── Event Listener Registrations ───────────────────────────────────────
  document.getElementById('profileFilterInput').addEventListener('input', (e) => {
    const filterText = e.target.value.toLowerCase();
    filterTable(filterText);
    chrome.storage.local.set({
      profileFilterValue: filterText,
      currentDisplayType: state.currentDisplayType
    });
  });
  document.getElementById("searchGroup").addEventListener("click", handleSearchGroup);
  document.getElementById("addToGroups").addEventListener("click", handleAddToGroups);
  document.getElementById("removeFromGroups").addEventListener("click", handleRemoveFromGroups);
  document.getElementById("checkGroups").addEventListener("click", handleCheckGroups);
  document.getElementById("checkGroupMembers").addEventListener("click", handleCheckGroupMembers);
  document.getElementById("checkCompliance").addEventListener("click", handleCheckCompliance);
  document.getElementById("downloadScript").addEventListener("click", handleDownloadScript);
  document.getElementById("appsAssignment").addEventListener("click", handleAppsAssignment);
  document.getElementById("pwshProfiles").addEventListener("click", handlePwshProfiles);
  document.getElementById("collectLogs").addEventListener("click", handleCollectLogs);
  document.getElementById("createGroup").addEventListener("click", handleCreateGroup); document.getElementById("groupResults").addEventListener("change", (event) => {
    if (event.target.type === "checkbox") {
      // Clear table selections when selecting checkboxes
      clearTableSelection();
      updateActionButtonsState();

      chrome.storage.local.get(['lastSearchResults'], (data) => {
        if (data.lastSearchResults) {
          const updated = data.lastSearchResults.map(group =>
            group.id === event.target.value ? { ...group, checked: event.target.checked } : group
          );
          chrome.storage.local.set({ lastSearchResults: updated });
        }
      });
    }
  });
  document.querySelector('th.sortable').addEventListener('click', function () {
    state.sortDirection = state.sortDirection === 'asc' ? 'desc' : 'asc';
    this.classList.toggle('asc');
    this.classList.toggle('desc');
    // Re-render the current table based on display type
    if (state.currentDisplayType === 'config') {
      chrome.storage.local.get(['lastConfigAssignments'], (data) => {
        if (data.lastConfigAssignments) updateConfigTable(data.lastConfigAssignments, false);
      });
    } else if (state.currentDisplayType === 'apps') {
      chrome.storage.local.get(['lastAppAssignments'], (data) => {
        if (data.lastAppAssignments) updateAppTable(data.lastAppAssignments, false);
      });
    } else if (state.currentDisplayType === 'compliance') {
      chrome.storage.local.get(['lastComplianceAssignments'], (data) => {
        if (data.lastComplianceAssignments) updateComplianceTable(data.lastComplianceAssignments, false);
      });
    } else if (state.currentDisplayType === 'pwsh') {
      chrome.storage.local.get(['lastPwshAssignments'], (data) => {
        if (data.lastPwshAssignments) updatePwshTable(data.lastPwshAssignments, false);
      });
    } else if (state.currentDisplayType === 'groupMembers') {
      chrome.storage.local.get(['lastGroupMembers'], (data) => {
        if (data.lastGroupMembers) updateGroupMembersTable(data.lastGroupMembers, false);
      });
    }
  });
  // Theme toggle button
  document.getElementById("theme-toggle").addEventListener("click", toggleTheme);

  // Pagination event listeners
  document.getElementById("prevPageBtn").addEventListener("click", () => {
    if (state.pagination.currentPage > 1) {
      goToPage(state.pagination.currentPage - 1);
    }
  });

  document.getElementById("nextPageBtn").addEventListener("click", () => {
    if (state.pagination.currentPage < state.pagination.totalPages) {
      goToPage(state.pagination.currentPage + 1);
    }
  });

  // Keyboard navigation for pagination
  document.addEventListener('keydown', (e) => {
    // Only handle pagination keys when table is visible and has data
    if (state.pagination.totalItems > 0) {
      if (e.key === 'ArrowLeft' && e.ctrlKey && state.pagination.currentPage > 1) {
        e.preventDefault();
        goToPage(state.pagination.currentPage - 1);
      } else if (e.key === 'ArrowRight' && e.ctrlKey && state.pagination.currentPage < state.pagination.totalPages) {
        e.preventDefault();
        goToPage(state.pagination.currentPage + 1);
      } else if (e.key === 'Home' && e.ctrlKey && state.pagination.currentPage > 1) {
        e.preventDefault();
        goToPage(1);
      } else if (e.key === 'End' && e.ctrlKey && state.pagination.currentPage < state.pagination.totalPages) {
        e.preventDefault();
        goToPage(state.pagination.totalPages);
      }
    }
  });

  // Target mode toggle buttons
  document.getElementById("deviceModeBtn").addEventListener("click", () => handleTargetModeToggle('device'));
  document.getElementById("userModeBtn").addEventListener("click", () => handleTargetModeToggle('user'));

  // ── Initial Restoration Calls ─────────────────────────────────────────
  restoreState();
  restoreFilterValue();
  initializeTheme();
});
