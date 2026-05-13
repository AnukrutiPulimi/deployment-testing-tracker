// Global variable to store selected profile
let selectedProfile = null;

// SharePoint site URL
const siteUrl = _spPageContextInfo ? _spPageContextInfo.webAbsoluteUrl : '';

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
  loadProfiles();
  setupEventListeners();
});

// Setup event listeners
function setupEventListeners() {
  const form = document.getElementById('createProfileForm');
  const dropdown = document.getElementById('profileDropdown');
  const deploymentForm = document.getElementById('createDeploymentForm');
  
  form.addEventListener('submit', function(e) {
    e.preventDefault();
    const profileName = document.getElementById('profileName').value.trim();
    const profileEmail = document.getElementById('profileEmail').value.trim();
    const profileRole = document.getElementById('profileRole').value.trim();
    
    if (profileName && profileEmail && profileRole) {
      createProfile(profileName, profileEmail, profileRole);
    }
  });

  dropdown.addEventListener('change', function() {
    if (this.value) {
      selectedProfile = JSON.parse(decodeURIComponent(this.value));
      showProfileSelected();
    } else {
      selectedProfile = null;
      hideProfileSection();
    }
  });

  deploymentForm.addEventListener('submit', function(e) {
    e.preventDefault();
    const weekStart = document.getElementById('deploymentWeekStart').value;
    const allocatedHours = parseInt(document.getElementById('deploymentAllocatedHours').value, 10);
    
    if (weekStart && allocatedHours > 0 && selectedProfile) {
      createDeployment(weekStart, allocatedHours);
    }
  });
}

// Show profile section, hide dashboard
function showProfileSection() {
  document.getElementById('profileSection').style.display = 'block';
  document.getElementById('dashboard').style.display = 'none';
}

// Hide profile section, show dashboard
function hideProfileSection() {
  document.getElementById('profileSection').style.display = 'none';
}

// Handle profile selection
function showProfileSelected() {
  document.getElementById('profileSection').style.display = 'none';
  document.getElementById('dashboard').style.display = 'block';
  document.getElementById('currentProfileName').textContent = `Logged in as: ${selectedProfile.Title} (${selectedProfile.Email})`;
  loadDeployments();
}

// Load profiles from SharePoint REST API
function loadProfiles() {
  if (!siteUrl) {
    showMessage('SharePoint context not available. Ensure page is hosted in SharePoint.', 'error');
    return;
  }

  const restUrl = `${siteUrl}/_api/web/lists/getbyTitle('Profiles')/items?$select=ID,Title,Email,Role`;
  
  fetch(restUrl, {
    method: 'GET',
    headers: {
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    }
  })
  .then(response => response.json())
  .then(data => {
    const dropdown = document.getElementById('profileDropdown');
    dropdown.innerHTML = '<option value="">-- Select a Profile --</option>';
    
    if (data.value && data.value.length > 0) {
      data.value.forEach(profile => {
        const option = document.createElement('option');
        const profileData = {
          ID: profile.ID,
          Title: profile.Title,
          Email: profile.Email,
          Role: profile.Role
        };
        option.value = encodeURIComponent(JSON.stringify(profileData));
        option.textContent = profile.Title;
        dropdown.appendChild(option);
      });
    }
  })
  .catch(error => {
    console.error('Error loading profiles:', error);
    showMessage('Error loading profiles. Check console for details.', 'error');
  });
}

// Load deployments filtered by ProfileEmail
function loadDeployments() {
  if (!selectedProfile || !siteUrl) {
    return;
  }

  const restUrl = `${siteUrl}/_api/web/lists/getbyTitle('Deployments')/items?$filter=ProfileEmail eq '${selectedProfile.Email}'&$orderby=WeekStart desc`;
  
  fetch(restUrl, {
    method: 'GET',
    headers: {
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    }
  })
  .then(response => response.json())
  .then(data => {
    const tbody = document.getElementById('deploymentsTableBody');
    const noData = document.getElementById('noDeployments');
    tbody.innerHTML = '';
    
    if (data.value && data.value.length > 0) {
      noData.style.display = 'none';
      data.value.forEach(deployment => {
        const row = createDeploymentRow(deployment);
        tbody.appendChild(row);
      });
    } else {
      noData.style.display = 'block';
    }
  })
  .catch(error => {
    console.error('Error loading deployments:', error);
    showDashboardMessage('Error loading deployments. Check console for details.', 'error');
  });
}

// Create deployment row in table
function createDeploymentRow(deployment) {
  const row = document.createElement('tr');
  
  const weekStart = new Date(deployment.WeekStart);
  const weekEnd = new Date(deployment.WeekEnd);
  const today = new Date();
  
  const usedHours = deployment.UsedHours || 0;
  const allocatedHours = deployment.AllocatedHours || 0;
  const remainingHours = deployment.RemainingHours !== undefined ? deployment.RemainingHours : (allocatedHours - usedHours);
  const isCompleted = deployment.IsCompleted || false;
  
  const isExpired = today > weekEnd && remainingHours > 0;
  
  row.innerHTML = `
    <td>${formatDate(weekStart)}</td>
    <td>${formatDate(weekEnd)}</td>
    <td>${allocatedHours}</td>
    <td>${usedHours}</td>
    <td>${remainingHours}</td>
    <td>${isCompleted ? 'Completed' : 'Active'}</td>
  `;
  
  if (isExpired) {
    row.classList.add('expired-warning');
  }
  
  return row;
}

// Format date to MM/DD/YYYY
function formatDate(date) {
  const d = new Date(date);
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const year = d.getFullYear();
  return `${month}/${day}/${year}`;
}

// Get request digest for POST operations
function getRequestDigest() {
  if (!siteUrl) {
    return Promise.reject('SharePoint context not available.');
  }

  const restUrl = `${siteUrl}/_api/contextinfo`;
  
  return fetch(restUrl, {
    method: 'POST',
    headers: {
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    }
  })
  .then(response => response.json())
  .then(data => data.FormDigestValue)
  .catch(error => {
    console.error('Error getting request digest:', error);
    throw error;
  });
}

// Create new profile in SharePoint
function createProfile(name, email, role) {
  getRequestDigest()
    .then(digest => {
      const restUrl = `${siteUrl}/_api/web/lists/getbyTitle('Profiles')/items`;
      const data = {
        Title: name,
        Email: email,
        Role: role
      };

      return fetch(restUrl, {
        method: 'POST',
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'X-RequestDigest': digest
        },
        body: JSON.stringify(data)
      });
    })
    .then(response => {
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      return response.json();
    })
    .then(profile => {
      showMessage(`Profile "${profile.Title}" created successfully!`, 'success');
      document.getElementById('createProfileForm').reset();
      selectedProfile = {
        ID: profile.ID,
        Title: profile.Title,
        Email: profile.Email,
        Role: profile.Role
      };
      loadProfiles();
    })
    .catch(error => {
      console.error('Error creating profile:', error);
      showMessage('Error creating profile. Check console for details.', 'error');
    });
}

// Create new deployment in SharePoint
function createDeployment(weekStart, allocatedHours) {
  const weekStartDate = new Date(weekStart);
  const weekEndDate = new Date(weekStartDate);
  weekEndDate.setDate(weekEndDate.getDate() + 7);
  
  getRequestDigest()
    .then(digest => {
      const restUrl = `${siteUrl}/_api/web/lists/getbyTitle('Deployments')/items`;
      const data = {
        ProfileEmail: selectedProfile.Email,
        WeekStart: weekStartDate.toISOString(),
        WeekEnd: weekEndDate.toISOString(),
        AllocatedHours: allocatedHours,
        UsedHours: 0,
        RemainingHours: allocatedHours,
        IsCompleted: false
      };

      return fetch(restUrl, {
        method: 'POST',
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'X-RequestDigest': digest
        },
        body: JSON.stringify(data)
      });
    })
    .then(response => {
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      return response.json();
    })
    .then(deployment => {
      showDashboardMessage(`Deployment created for week of ${formatDate(deployment.WeekStart)}!`, 'success');
      document.getElementById('createDeploymentForm').reset();
      loadDeployments();
    })
    .catch(error => {
      console.error('Error creating deployment:', error);
      showDashboardMessage('Error creating deployment. Check console for details.', 'error');
    });
}

// Show message to user (profile section)
function showMessage(text, type) {
  const messageDiv = document.getElementById('message');
  messageDiv.textContent = text;
  messageDiv.className = `message ${type}`;
  
  if (type === 'success') {
    setTimeout(() => {
      messageDiv.className = 'message';
    }, 4000);
  }
}

// Show message to user (dashboard section)
function showDashboardMessage(text, type) {
  const messageDiv = document.getElementById('dashboardMessage');
  messageDiv.textContent = text;
  messageDiv.className = `message ${type}`;
  
  if (type === 'success') {
    setTimeout(() => {
      messageDiv.className = 'message';
    }, 4000);
  }
}
