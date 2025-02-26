// -------------------------------
// Helper Functions
// -------------------------------

// Get the Dataverse Base URL from the active tab.
async function getDataverseUrl() {
  return new Promise((resolve, reject) => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      const activeTabUrl = new URL(tabs[0].url);
      if (activeTabUrl.hostname.endsWith(".dynamics.com")) {
        resolve(`${activeTabUrl.protocol}//${activeTabUrl.hostname}`);
      } else {
        reject(new Error("Not on a Dataverse page. Please open your Dataverse instance."));
      }
    });
  });
}

// Generic Fetch API helper
async function fetchDataverseData(apiUrl) {
  try {
    const response = await fetch(apiUrl, {
      method: "GET",
      headers: {
        "OData-Version": "4.0",
        "OData-MaxVersion": "4.0",
        Accept: "application/json",
        "Content-Type": "application/json"
      },
      credentials: "include"
    });
    if (!response.ok) {
      throw new Error(`Failed to fetch data: ${response.status} ${response.statusText}`);
    }
    return response.json();
  } catch (error) {
    console.log("Error fetching data:" + error);
    return null;
  }
}

// -------------------------------
// Functions for the "User Roles" Tab
// -------------------------------

async function getUserSecurityRoles(fullname) {
  const baseUrl = await getDataverseUrl();
  showLoading(true);

  // Query for direct role assignments via systemuserroles.
  const fetchUserRoleXml = `
      <fetch distinct="true">
        <entity name="role">
          <attribute name="name" />
          <link-entity name="systemuserroles" from="roleid" to="roleid" intersect="true">
            <link-entity name="systemuser" from="systemuserid" to="systemuserid" intersect="true">
              <filter>
                <condition attribute="fullname" operator="eq" value="${fullname}" />
              </filter>
            </link-entity>
          </link-entity>
        </entity>
      </fetch>`;

  // Query for team-based role assignments via teamroles.
  const fetchTeamsRoleXml = `
      <fetch distinct="true">
        <entity name="role">
          <attribute name="name" />
          <link-entity name="teamroles" from="roleid" to="roleid" intersect="true">
            <link-entity name="team" from="teamid" to="teamid" intersect="true">
              <attribute name="name" />
              <link-entity name="teammembership" from="teamid" to="teamid" intersect="true">
                <link-entity name="systemuser" from="systemuserid" to="systemuserid" intersect="true">
                  <filter>
                    <condition attribute="fullname" operator="eq" value="${fullname}" />
                  </filter>
                </link-entity>
              </link-entity>
            </link-entity>
          </link-entity>
        </entity>
      </fetch>`;

  try {
    const userRolesData = await fetchDataverseData(
      `${baseUrl}/api/data/v9.1/roles?fetchXml=${encodeURIComponent(fetchUserRoleXml)}`
    );
    const teamRolesData = await fetchDataverseData(
      `${baseUrl}/api/data/v9.1/roles?fetchXml=${encodeURIComponent(fetchTeamsRoleXml)}`
    );

    // Map the results into a unified list.
    const userRoles =
      userRolesData?.value?.map((role) => ({
        name: role.name,
        teamName: "N/A",
        source: "Direct"
      })) || [];
    const teamRoles =
      teamRolesData?.value?.map((role) => ({
        name: role.name,
        teamName: role["team.name"] || "Unknown",
        source: "Team"
      })) || [];

    const roleMap = new Map();
    [...userRoles, ...teamRoles].forEach((role) => {
      if (!roleMap.has(role.name)) {
        roleMap.set(role.name, role);
      } else {
        // Mark as assigned both directly and via team.
        roleMap.set(role.name, {
          name: role.name,
          teamName: role.teamName,
          source: "Direct & Team"
        });
      }
    });
    const uniqueRoles = Array.from(roleMap.values());

    if (uniqueRoles.length === 0) {
      displayRolesTable(uniqueRoles);
      displayMessage(`No security roles found for user '${fullname}'.`, true);
    } else {
      displayRolesTable(uniqueRoles);
      displayMessage(`Found ${uniqueRoles.length} role(s) for '${fullname}'.`);
    }
  } catch (error) {
    console.log("Error fetching user roles:" + error);
    displayMessage("Failed to fetch user roles. Please try again.", true);
  } finally {
    showLoading(false);
  }
}

function displayRolesTable(roles) {
  debugger;
  const tableContainer = document.getElementById("rolesTableContainer");
  if (!tableContainer) return;

  // Clear the container before rendering new content.
  tableContainer.innerHTML = "";

  // If no roles are returned, clear the table container and optionally display a message.
  if (!roles || roles.length === 0) {
    tableContainer.innerHTML = "<p>No security roles found.</p>";
    return;
  }

  let tableHtml = `
      <table border="1">
        <thead>
          <tr>
            <th>Role Name</th>
            <th>Team Name</th>
            <th>Source</th>
          </tr>
        </thead>
        <tbody>`;
  roles.forEach((role) => {
    tableHtml += `
        <tr>
          <td>${role.name}</td>
          <td>${role.teamName || "N/A"}</td>
          <td>${role.source}</td>
        </tr>`;
  });
  tableHtml += `</tbody></table>`;
  tableContainer.innerHTML = tableHtml;
}

function showLoading(isLoading) {
  const searchButton = document.getElementById("searchUserButton");
  if (isLoading) {
    searchButton.disabled = true;
    searchButton.textContent = "Searching...";
  } else {
    searchButton.disabled = false;
    searchButton.textContent = "Search User";
  }
}

function displayMessage(message, isError = false) {
  const feedbackDiv = document.getElementById("feedbackMessage");
  feedbackDiv.innerHTML = `<p style="color: ${isError ? "red" : "green"}; margin: 5px 0;">${message}</p>`;
  setTimeout(() => {
    feedbackDiv.innerHTML = "";
  }, 5000);
}

// -------------------------------
// Functions for the "Security Role Members" Tab
// -------------------------------

// 1. Get the list of security roles.
async function getSecurityRoles() {
  const baseUrl = await getDataverseUrl();
  const fetchXml = `
      <fetch distinct="true">
        <entity name="role">
          <attribute name="name" />
          <attribute name="roleid" />
          <order attribute="name" descending="false" />
        </entity>
      </fetch>`;
  const url = `${baseUrl}/api/data/v9.1/roles?fetchXml=${encodeURIComponent(fetchXml)}`;
  return await fetchDataverseData(url);
}

// 2. Populate the security roles dropdown.
function populateSecurityRolesDropdown(data) {
  const dropdown = document.getElementById("securityRoles");
  dropdown.innerHTML = ""; // Clear any existing options.
  if (data && data.value && data.value.length > 0) {
    data.value.forEach((role) => {
      const option = document.createElement("option");
      option.value = role.roleid;
      option.text = role.name;
      dropdown.appendChild(option);
    });
  } else {
    const option = document.createElement("option");
    option.value = "";
    option.text = "No security roles found";
    dropdown.appendChild(option);
  }
}

// 3. Retrieve users assigned to a given security role (directly or via teams).
async function getUsersBySecurityRole(roleid) {
  const baseUrl = await getDataverseUrl();

  // Fetch users with direct role assignments
  const fetchDirectXml = `
    <fetch distinct="true">
      <entity name="systemuser">
        <attribute name="fullname" />
        <attribute name="systemuserid" />
        <link-entity name="systemuserroles" from="systemuserid" to="systemuserid">
          <link-entity name="role" from="roleid" to="roleid">
            <filter>
              <condition attribute="roleid" operator="eq" value="${roleid}" />
            </filter>
          </link-entity>
        </link-entity>
      </entity>
    </fetch>`;

  // Fetch users with role assignments through teams
  const fetchTeamXml = `
    <fetch distinct="true">
      <entity name="systemuser">
        <attribute name="fullname" />
        <attribute name="systemuserid" />
        <link-entity name="teammembership" from="systemuserid" to="systemuserid">
          <link-entity name="teamroles" from="teamid" to="teamid">
            <link-entity name="role" from="roleid" to="roleid">
              <filter>
                <condition attribute="roleid" operator="eq" value="${roleid}" />
              </filter>
            </link-entity>
            <link-entity name="team" from="teamid" to="teamid" alias="team">
              <attribute name="name" />
            </link-entity>
          </link-entity>
        </link-entity>
      </entity>
    </fetch>`;

  const directUsersData = await fetchDataverseData(`${baseUrl}/api/data/v9.1/systemusers?fetchXml=${encodeURIComponent(fetchDirectXml)}`);
  const teamUsersData = await fetchDataverseData(`${baseUrl}/api/data/v9.1/systemusers?fetchXml=${encodeURIComponent(fetchTeamXml)}`);

  // Combine results and add source information
  const usersMap = new Map();

  if (directUsersData && directUsersData.value) {
    directUsersData.value.forEach((user) => {
      usersMap.set(user.systemuserid, {
        fullname: user.fullname,
        userId: user.systemuserid,
        source: "Direct",
        teamName: "N/A"
      });
    });
  }

  if (teamUsersData && teamUsersData.value) {
    teamUsersData.value.forEach((user) => {
      if (!usersMap.has(user.systemuserid)) {
        usersMap.set(user.systemuserid, {
          fullname: user.fullname,
          userId: user.systemuserid,
          source: "Team",
          teamName: user["team.name"] || "Unknown"
        });
      }
    });
  }

  return Array.from(usersMap.values());
}


// 4. Display the users in a table.
function displayRoleUsersTable(users) {
  const container = document.getElementById("roleUsersTableContainer");
  let tableHtml = `
    <table border="1">
      <thead>
        <tr>
          <th>Full Name</th>
          <th>User ID</th>
          <th>Source</th>
          <th>Team Name</th>
        </tr>
      </thead>
      <tbody>`;
  if (users.length > 0) {
    users.forEach((user) => {
      tableHtml += `
        <tr>
          <td>${user.fullname}</td>
          <td>${user.userId}</td>
          <td>${user.source}</td>
          <td>${user.teamName}</td>
        </tr>`;
    });
  } else {
    tableHtml += `<tr><td colspan="4">No users found for this security role.</td></tr>`;
  }
  tableHtml += `</tbody></table>`;
  container.innerHTML = tableHtml;
}


function displayRoleUsersFeedback(message, isError = false) {
  const feedbackDiv = document.getElementById("roleUsersFeedbackMessage");
  feedbackDiv.innerHTML = `<p style="color: ${isError ? "red" : "green"}; margin: 5px 0;">${message}</p>`;
  setTimeout(() => {
    feedbackDiv.innerHTML = "";
  }, 5000);
}

// -------------------------------
// Event Listeners
// -------------------------------

// For fetching user roles by fullname.
/*
document.getElementById("fetchRoles").addEventListener("click", async function () {
  const fullname = document.getElementById("fullname").value.trim();
  if (fullname) {
    await getUserSecurityRoles(fullname);
  } else {
    alert("Please enter the fullname.");
  }
});
*/

// For applying the selected security role.
document.getElementById("applyRole").addEventListener("click", async function () {
  const dropdown = document.getElementById("securityRoles");
  const selectedRoleId = dropdown.value;
  if (selectedRoleId) {
    displayRoleUsersFeedback("Loading users...", false);
    const users = await getUsersBySecurityRole(selectedRoleId);
    displayRoleUsersTable(users);
    displayRoleUsersFeedback(`Found ${users.length} user(s).`, false);
  } else {
    alert("Please select a security role.");
  }
});

// -------------------------------
// Tab Switching Logic
// -------------------------------
document.querySelectorAll(".tab-button").forEach((btn) => {
  btn.addEventListener("click", function () {
    // Remove active class from all tab buttons and contents.
    document.querySelectorAll(".tab-button").forEach((b) => b.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach((tab) => tab.classList.remove("active"));
    // Activate the clicked tab button and corresponding tab content.
    this.classList.add("active");
    const tabId = this.getAttribute("data-tab");
    document.getElementById(tabId).classList.add("active");
  });
});

// -------------------------------
// On Popup Load: Populate the Security Roles Dropdown
// -------------------------------
document.addEventListener("DOMContentLoaded", async function () {
  const rolesData = await getSecurityRoles();
  populateSecurityRolesDropdown(rolesData);
});


// ---------------------------------
// Function to search for users by partial fullname
// ---------------------------------
async function searchUsers(searchTerm) {
  const baseUrl = await getDataverseUrl();
  const fetchXml = `
    <fetch distinct="true">
      <entity name="systemuser">
        <attribute name="fullname" />
        <attribute name="systemuserid" />
        <filter type="and">
          <condition attribute="fullname" operator="like" value="%${searchTerm}%" />
        </filter>
      </entity>
    </fetch>`;
  const url = `${baseUrl}/api/data/v9.1/systemusers?fetchXml=${encodeURIComponent(fetchXml)}`;
  return await fetchDataverseData(url);
}

// ---------------------------------
// Function to populate the user search dropdown
// ---------------------------------
function populateUserDropdown(usersData) {
  const dropdown = document.getElementById("userSearchDropdown");
  // Clear previous entries
  dropdown.innerHTML = "";
  if (usersData && usersData.value && usersData.value.length > 0) {
    // Add a default option
    const defaultOption = document.createElement("option");
    defaultOption.value = "";
    defaultOption.text = "Select a user";
    dropdown.appendChild(defaultOption);

    usersData.value.forEach((user) => {
      const option = document.createElement("option");
      // Save the fullname in the value so it can be used to fetch roles
      option.value = user.fullname;
      option.text = user.fullname;
      dropdown.appendChild(option);
    });
  } else {
    const option = document.createElement("option");
    option.value = "";
    option.text = "No users found";
    dropdown.appendChild(option);
  }
}

// ---------------------------------
// Event Listener for the user search button
// ---------------------------------
document.getElementById("searchUserButton").addEventListener("click", async function () {
  const searchTerm = document.getElementById("userSearchInput").value.trim();
  if (searchTerm) {
    // Optionally, clear previous search results and show a temporary loading state
    const dropdown = document.getElementById("userSearchDropdown");
    dropdown.innerHTML = "";
    const loadingOption = document.createElement("option");
    loadingOption.text = "Loading...";
    dropdown.appendChild(loadingOption);

    // Perform the search
    const usersData = await searchUsers(searchTerm);
    populateUserDropdown(usersData);
  } else {
    alert("Please enter a search term.");
  }
});

// ---------------------------------
// Event Listener to handle selection from the dropdown
// ---------------------------------
// When a user is selected, use their full name to retrieve security roles
document.getElementById("userSearchDropdown").addEventListener("change", async function () {
  const selectedFullName = this.value;
  if (selectedFullName) {
    // Use the existing function to fetch and display user roles by fullname
    await getUserSecurityRoles(selectedFullName);
  }
});
