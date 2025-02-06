
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
        console.log("Error fetching data:", error);
        return null;
    }
}

// Fetch User Security Roles
async function getUserSecurityRoles(fullname) {
    const baseUrl = await getDataverseUrl();

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

    // Fetch user roles (direct assignments)
    const userRolesData = await fetchDataverseData(`${baseUrl}/api/data/v9.1/roles?fetchXml=${encodeURIComponent(fetchUserRoleXml)}`);

    // Fetch team roles (roles assigned via teams)
    const teamRolesData = await fetchDataverseData(`${baseUrl}/api/data/v9.1/roles?fetchXml=${encodeURIComponent(fetchTeamsRoleXml)}`);

    // Extract roles, team names, and their source (Direct or Team)
    const userRoles = userRolesData?.value?.map(role => ({ name: role.name, teamName: "N/A", source: "Direct" })) || [];
    const teamRoles = teamRolesData?.value?.map(role => ({ name: role.name, teamName: role["team.name"] || "Unknown", source: "Team" })) || [];

    // Combine results and remove duplicates (keeping source information)
    const uniqueRoles = [];
    const roleMap = new Map();

    [...userRoles, ...teamRoles].forEach(role => {
        if (!roleMap.has(role.name)) {
            roleMap.set(role.name, role);
        } else {
            // If the same role exists in both sources, update source to "Direct & Team"
            roleMap.set(role.name, { name: role.name, teamName: role.teamName, source: "Direct & Team" });
        }
    });

    uniqueRoles.push(...roleMap.values());

    if (uniqueRoles.length === 0) {
        alert(`No security roles found for user '${fullname}'.`);
        return;
    }

    // Display roles with their source (Direct / Team / Both)
    displayRolesTable(uniqueRoles);
}


function displayRolesTable(roles) {
    const tableContainer = document.getElementById("rolesTableContainer");
    if (!tableContainer) return;

    if (roles.length === 0) {
        tableContainer.innerHTML = "<p>No roles found.</p>";
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



// Event Listener for Fetching User Roles
document.getElementById("fetchRoles").addEventListener("click", function () {
    const fullname = document.getElementById("fullname").value.trim();
    if (fullname) {
        getUserSecurityRoles(fullname);
    } else {
        alert("Please enter the fullname.");
    }
});

// Toggle User Role Section
function toggleUserRoleSection() {
    const section = document.getElementById("userRoleSection");
    section.style.display = section.style.display === "none" ? "block" : "none";
}

// Tab Navigation Logic
document.addEventListener("DOMContentLoaded", function () {
    const tabButtons = document.querySelectorAll(".tab-button");
    const tabContents = document.querySelectorAll(".tab-content");

    tabButtons.forEach(button => {
        button.addEventListener("click", function () {
            tabButtons.forEach(btn => btn.classList.remove("active"));
            tabContents.forEach(content => content.classList.remove("active"));

            const tabId = this.getAttribute("data-tab");
            this.classList.add("active");
            document.getElementById(tabId).classList.add("active");
        });
    });
});