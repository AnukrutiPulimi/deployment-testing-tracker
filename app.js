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
      showMessage(`Selected profile: ${selectedProfile.Title}`, 'success');
    } else {
      selectedProfile = null;
    }
  });
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

// Show message to user
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
