
document.addEventListener('DOMContentLoaded', async () => {
    checkAuthState();
    loadProfile();

    const findFriendBtn = document.getElementById('findFriendBtn');
    if(findFriendBtn) {
        findFriendBtn.addEventListener('click', findFriends);
    }
});

function checkAuthState() {
    const token = localStorage.getItem('token');
    const username = localStorage.getItem('username');
    if (token && username) {
        updateUIForAuth(username);
    } else {
        window.location.href = 'index.html';
    }
}

function updateUIForAuth(username) {
    const userInfo = document.getElementById('userInfo');
    const profileLink = document.getElementById('profileLink');
    const loginBtn = document.getElementById('loginBtn');
    const registerBtn = document.getElementById('registerBtn');

    if (userInfo && profileLink && loginBtn && registerBtn) {
        loginBtn.style.display = 'none';
        registerBtn.style.display = 'none';
        userInfo.style.display = 'flex';
        profileLink.textContent = username;
        profileLink.href = '#';
    }
}

async function loadProfile() {
    const token = localStorage.getItem('token');
    if (!token) return;

    try {
        const response = await fetch('/api/profile', {
            headers: { 'Authorization': token }
        });

        if (response.ok) {
            const profile = await response.json();
            displayProfile(profile);
        } else {
            console.error('Failed to load profile');
        }
    } catch (error) {
        console.error('Error loading profile:', error);
    }
}

function displayProfile(profile) {
    const profileContent = document.getElementById('profileContent');
    const friendsList = document.getElementById('friendsList');

    if (profileContent) {
        profileContent.innerHTML = `<p>Username: ${profile.username}</p>`;
    }

    if (friendsList) {
        friendsList.innerHTML = '';
        profile.friends.forEach(friend => {
            const li = document.createElement('li');
            li.textContent = friend;
            friendsList.appendChild(li);
        });
    }
}

async function findFriends() {
    const searchInput = document.getElementById('findFriendInput');
    const searchTerm = searchInput.value;
    const searchResults = document.getElementById('searchResults');

    if (!searchTerm) {
        searchResults.innerHTML = '<p>Please enter a username to search.</p>';
        return;
    }

    try {
        const response = await fetch(`/api/users/search?username=${searchTerm}`);
        if(response.ok) {
            const users = await response.json();
            let html = '<ul>';
            users.forEach(user => {
                html += `<li>${user.username}</li>`;
            });
            html += '</ul>';
            searchResults.innerHTML = html;
        } else {
            searchResults.innerHTML = '<p>No users found.</p>';
        }
    } catch (error) {
        console.error('Error searching for users:', error);
        searchResults.innerHTML = '<p>An error occurred while searching.</p>';
    }
}
