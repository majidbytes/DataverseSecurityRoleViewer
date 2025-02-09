// Get Dataverse Base URL
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

// Generic Fetch API Helper
async function fetchDataverseData(apiUrl) {
    try {
        const response = await fetch(apiUrl, {
            method: "GET",
            headers: {
                "OData-Version": "4.0",
                "OData-MaxVersion": "4.0",
                "Accept": "application/json",
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

// Fetch User Security Roles
async function getUserSecurityRoles(fullname) {
    const baseUrl = await getDataverseUrl();
    showLoading(true);

    const fetchUserRoleXml = `
    <fetch distinct="true">
      <entity name="role">
        <attribute name="name" />
        <link-entity name="systemuserroles" from="roleid" to="roleid" alias="role" intersect="true">
          <link-entity name="systemuser" from="systemuserid" to="systemuserid" alias="user" intersect="true">
            <filter>
              <condition attribute="fullname" operator="eq" value="${fullname}" />
            </filter>
          </link-entity>
        </link-entity>
      </entity>
    </fetch>`;

    const fetchTeamsRoleXml = `
    <fetch distinct="true">
      <entity name="role">
        <attribute name="name" />
        <link-entity name="teamroles" from="roleid" to="roleid" alias="teamrole" intersect="true">
          <link-entity name="team" from="teamid" to="teamid" alias="team" intersect="true">
            <attribute name="name" />
            <link-entity name="teammembership" from="teamid" to="teamid" intersect="true">
              <link-entity name="systemuser" from="systemuserid" to="systemuserid" alias="teammember" intersect="true">
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
        const userRolesData = await fetchDataverseData(`${baseUrl}/api/data/v9.1/roles?fetchXml=${encodeURIComponent(fetchUserRoleXml)}`);
        const teamRolesData = await fetchDataverseData(`${baseUrl}/api/data/v9.1/roles?fetchXml=${encodeURIComponent(fetchTeamsRoleXml)}`);

        const userRoles = userRolesData?.value?.map(role => ({ name: role.name, teamName: "N/A", source: "Direct" })) || [];
        const teamRoles = teamRolesData?.value?.map(role => ({ name: role.name, teamName: role["team.name"] || "Unknown", source: "Team" })) || [];

        const uniqueRoles = [];
        const roleMap = new Map();

        [...userRoles, ...teamRoles].forEach(role => {
            if (!roleMap.has(role.name)) {
                roleMap.set(role.name, role);
            } else {
                roleMap.set(role.name, { name: role.name, teamName: role.teamName, source: "Direct & Team" });
            }
        });

        uniqueRoles.push(...roleMap.values());

        if (uniqueRoles.length === 0) {
            displayMessage(`No security roles found for user '${fullname}'.`, true);
        } else {
            displayRolesTable(uniqueRoles);
            displayMessage(`Found ${uniqueRoles.length} roles for '${fullname}'.`);
        }
    } catch (error) {
        console.log("Error fetching user roles:" + error);
        displayMessage("Failed to fetch user roles. Please try again.", true);
    } finally {
        showLoading(false);
    }
}

// Display roles in a table
function displayRolesTable(roles) {
    const tableContainer = document.getElementById("rolesTableContainer");
    if (!tableContainer) return;

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

    roles.forEach(role => {
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

// Show or hide loading indicator
function showLoading(isLoading) {
    const fetchButton = document.getElementById("fetchRoles");
    if (isLoading) {
        fetchButton.disabled = true;
        fetchButton.textContent = "Loading...";
    } else {
        fetchButton.disabled = false;
        fetchButton.textContent = "Get User Roles";
    }
}

// Display feedback message
function displayMessage(message, isError = false) {
    const feedbackDiv = document.getElementById("feedbackMessage");
    feedbackDiv.innerHTML = `
        <p style="color: ${isError ? 'red' : 'green'}; margin: 5px 0;">
            ${message}
        </p>
    `;
    // Automatically clear the message after 5 seconds
    setTimeout(() => {
        feedbackDiv.innerHTML = "";
    }, 5000);
}


// Event Listener for Fetching User Roles
document.getElementById("fetchRoles").addEventListener("click", async function () {
    const fullname = document.getElementById("fullname").value.trim();
    if (fullname) {
        await getUserSecurityRoles(fullname);
    } else {
        alert("Please enter the fullname.");
    }
});
