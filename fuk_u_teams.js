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
let forcedAvailability = "";
let isScriptLoaded = false;
let dompurify_policy; // trusted-types policy usable to create the selector


// Show the note when people message you
function pinned(s) {
    return s + "<pinnednote></pinnednote>"
}

const statusMap = {
    "Script Off": {"availability":"", "note": ""},
    "Available": {"availability":"Available", "note": ""},
    "Busy": {"availability":"Busy", "note": ""},
    "DoNotDisturb": {"availability":"DoNotDisturb", "note": ""},
    "BeRightBack": {"availability":"BeRightBack", "note": ""},
    "Away": {"availability":"Away", "note": ""},
    "Offline": {"availability":"Offline", "note": ""}
};

function getAuthToken() {
    for(const i in localStorage) {
        if(i.startsWith("ts.") && i.endsWith("cache.token.https://presence.teams.microsoft.com/")) {
            return JSON.parse(localStorage[i]).token;
        }
    }
}

function forceStatus() {
    if(!forcedAvailability) return;
    console.log(`fuk u teams, I'm ${forcedAvailability}`);
    fetch("https://presence.teams.microsoft.com/v1/me/forceavailability/", {
        "headers": {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${getAuthToken()}`
        },
        "body": `{"availability":"${forcedAvailability}"}`,
        "method": "PUT"
    })
    .then((resp) => console.log(`(forceStatus) Got fuked: ${resp.status}`))
    .catch((err) => console.error("Unable to set the new status:", err));
}

function resetStatus() {
    console.log(`reset status`);
    fetch("https://presence.teams.microsoft.com/v1/me/forceavailability/", {
        "headers": {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${getAuthToken()}`
        },
        "method": "PUT"
    })
    .then((resp) => console.log(`(resetStatus): ${resp.status}`))
    .catch((err) => console.error("Unable to reset the status:", err));
}

function checkIfSelectorChanged() {
    // the teams security policy do not permit the usage of custom code in "onchange" event
    let node = document.getElementById(selectorId);
    let newStatus = statusMap[node.value];
    if(newStatus && newStatus.availability != forcedAvailability) {
        publishNote(newStatus.note);
        forcedAvailability = newStatus.availability;
        if(forcedAvailability) {
            forceStatus();
        } else {
            resetStatus();
        }
    }
}

function createSelector() {
    let node = document.createElement("div");
    node.style.position = 'absolute';
    node.style.top = '15px';
    node.style.right = '100px';
    node.style.zIndex = '999999';

    let options = "";
    for(let status in statusMap){
        options += `<option value="${status}">${status}</option>`;
    }
    node.innerHTML = dompurify_policy.createHTML(`
    <select id="${selectorId}" style="color: black;">
        ${options}
    </select>`);
    return node;
}

function publishNote(note) {
    console.log(`fuk u teams, tells the others "${note}"`);
    fetch("https://presence.teams.microsoft.com/v1/me/publishnote", {
        "headers": {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${getAuthToken()}`
        },
        "body": `{"message": "${note}", "expiry": "9999-12-30T23:00:00.000Z"}`,
        "method": "PUT"
    })
    .then((resp) => console.log(`(publishNote) Got fuked: ${resp.status}`))
    .catch((err) => console.error("Unable to publish the new note:", err));
}

// catch the dompurify policy to reuse it to create the selector
{
    const createPolicy = trustedTypes['createPolicy'];
    trustedTypes['createPolicy'] = function ovr(name, opts) {
        const p = createPolicy.call(trustedTypes, name, opts);
        if (name == '@msteams/frameworks-loader#dompurify') {
            dompurify_policy = p;
        }
        return p;
    };
}

window.addEventListener('DOMContentLoaded', () => {
    document.body.appendChild(createSelector());
    setInterval(checkIfSelectorChanged, 1000); // 1s
    setInterval(forceStatus, 15 * 1000); // 15s
}, { once: true });
