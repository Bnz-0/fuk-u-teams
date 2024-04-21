// ==UserScript==
// @name        fuk u teams
// @version     3
// @description Script for keeping status perpetually in Web version of Teams
// @run-at      document-start
// @grant       none
// @match       *://*.teams.microsoft.com/*
// @author      Bnz-0, tetroxid (https://www.reddit.com/user/tetroxid/)
// ==/UserScript==


const selectorId = "FUK_U_TEAMS_SELECTOR";
let forcedStatus = "Script Off"; // name from the statusMap
let isScriptLoaded = false;
let dompurify_policy; // trusted-types policy usable to create the selector


// Show the note when people message you
function pinned(s) {
    return s + "<pinnednote></pinnednote>"
}

const RESET = "";
const AVAILABLE = { "availability": "Available" };
const BUSY = { "availability": "Busy" };
const DO_NOT_DISTURB = { "availability": "DoNotDisturb" };
const BE_RIGHT_BACK = { "availability": "BeRightBack" };
const AWAY = { "availability": "Away" };
const OFFLINE = { "availability": "Offline", "activity": "OffWork" };

const statusMap = {
    "Script Off": { "status": RESET },
    "Available": { "status": AVAILABLE },
    "Busy": { "status": BUSY },
    "DoNotDisturb": { "status": DO_NOT_DISTURB },
    "BeRightBack": { "status": BE_RIGHT_BACK },
    "Away": { "status": AWAY },
    "Offline": { "status": OFFLINE }
};


function getAuthToken() {
    for (const i in localStorage) {
        if (i.indexOf("https://presence.teams.microsoft.com//.default") >= 0) {
            return JSON.parse(localStorage[i]).secret;
        }
    }
}

function setStatus(status) {
    status = status ? JSON.stringify(status) : undefined;
    fetch("https://presence.teams.microsoft.com/v1/me/forceavailability/", {
        "headers": {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${getAuthToken()}`
        },
        "body": status,
        "method": "PUT"
    })
        .catch((err) => console.error("[fuk-u-teams] Unable to set the new status:", err));
}

function publishNote(note) {
    fetch("https://presence.teams.microsoft.com/v1/me/publishnote", {
        "headers": {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${getAuthToken()}`
        },
        "body": `{"message": "${note}", "expiry": "9999-12-30T23:00:00.000Z"}`,
        "method": "PUT"
    })
        .catch((err) => console.error("[fuk-u-teams] Unable to publish the new note:", err));
}

function forceStatus() {
    let selected_status = statusMap[forcedStatus];
    if (selected_status && selected_status.status != RESET) {
        setStatus(selected_status.status);
    }
}

function checkIfSelectorChanged() {
    // the teams security policy do not permit the usage of custom code in "onchange" event
    let node = document.getElementById(selectorId);
    let newStatus = node.value;
    if (newStatus != forcedStatus) {
        let old_status = statusMap[forcedStatus];
        let selected_status = statusMap[newStatus];
        if (selected_status.status.availability != old_status.status.availability) {
            setStatus(selected_status.status);
        }
        if (selected_status.note != old_status.note) {
            publishNote(selected_status.note || "");
        }
        forcedStatus = newStatus;
    }
}

function createSelector() {
    let node = document.createElement("div");
    node.style.position = 'absolute';
    node.style.top = '15px';
    node.style.right = '100px';
    node.style.zIndex = '999999';

    let options = "";
    for (let status in statusMap) {
        options += `<option value="${status}">${status}</option>`;
    }
    node.innerHTML = dompurify_policy.createHTML(`
    <select id="${selectorId}" style="color: black;">
        ${options}
    </select>`);
    return node;
}


// catch the dompurify policy to reuse it to create the selector
{
    const createPolicy = trustedTypes['createPolicy'];
    trustedTypes['createPolicy'] = function ovr(name, opts) {
        const p = createPolicy.call(trustedTypes, name, opts);
        if (name == 'dompurify') {
            dompurify_policy = p;
        }
        return p;
    };
}

function init() {
	if(dompurify_policy) {
	    document.body.appendChild(createSelector());
	    setInterval(checkIfSelectorChanged, 1000); // 1s
	    setInterval(forceStatus, 15 * 1000); // 15s
	} else {
		setTimeout(init, 1000);
	}
}

window.addEventListener('DOMContentLoaded', () => {
	init();
}, { once: true });
