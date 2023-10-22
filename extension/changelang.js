var botDisableGB = false;

function openSpellingMenu() {
    var collection = document.getElementsByTagName("button");
    for (let i = 0; i < collection.length; i++) {
        if (collection[i].getAttribute("aria-label") == "Spelling Show More Options") {
            collection[i].click();
        }
    }
}

function clickSetProofingLanguage() {
    var collection = document.getElementsByTagName("span");
    for (let i = 0; i < collection.length; i++) {
        if (collection[i].innerHTML == "Set Proofing Language...") {
            collection[i].click();
        }
    }
}

function findUKButton() {
    let foundSPL = false;
    var collection = document.getElementsByTagName("option");
    for (let i = 0; i < collection.length; i++) {
        if (collection[i].innerHTML == "English (U.K.)") {
            collection[i].selected = true;
            foundSPL = true;
        }
    }
    return foundSPL;
}
function clickOK() {
    var collection = document.getElementsByTagName("button");
    for (let i = 0; i < collection.length; i++) {
        if (collection[i].id == "WACDialogActionButton" && collection[i].innerHTML == "OK") {
            collection[i].click();
        }
    }
}

function getCurrentPageID(){
    var collection = document.getElementsByClassName("pageNode")
    for (let i = 0; i < collection.length; i++) {
        // if the 2nd of 3 child elements classname contains "elected"
        if (collection[i].children[1].className.indexOf("elected") > -1) {
            return collection[i].id;
        }
    }
}

function setSpellingLang(){
    // check if the bot is disabled
    if(!botDisableGB){
        openSpellingMenu();
        // wait for the menu to open
        setTimeout(function() {
            clickSetProofingLanguage();
            // wait for the menu to open
            setTimeout(function() {
                findUKButton();
                // wait for the menu to open
                setTimeout(function() {
                    clickOK();
                }, 40);
            }, 40);
        }, 40);
    }
}

// if window.location.href contains officeapps.live.com
if(window.location.href.indexOf("officeapps.live.com") > -1) {
    // if window.location.href contains onenoteframe
    if(window.location.href.indexOf("onenoteframe") > -1) {
        // give it a second to load
        setTimeout(function() {
            // set the spelling language
            setSpellingLang();
            // set the last page id
            window.lastPageID = getCurrentPageID();
        }, 1000);
        setInterval(function() {
            // if the page has changed
            if(getCurrentPageID() != window.lastPageID){
                // give it a second to load
                setTimeout(function() {
                    // set the spelling language
                    setSpellingLang();
                    // set the last page id
                    window.lastPageID = getCurrentPageID();
                }, 1000);
            }
        }, 2000);
        // create a button that is in the top left corner with a british flag on it that when clicked sets the spelling language
        var btn = document.createElement("BUTTON");
        btn.innerHTML = "🇬🇧";
        btn.style = "position: absolute; top: 0px; left: 0px; z-index: 9999; font-size: 20px;";
        btn.onclick = function() {setSpellingLang()};
        document.body.appendChild(btn);
        // create a second button called "disable GB bot" thats in the top left corner offset by 100px that when clicked disables the bot
        var btn2 = document.createElement("BUTTON");
        btn2.innerHTML = "turn off auto GB spelling bot";
        btn2.style = "position: absolute; top: 0px; left: 100px; z-index: 9999; font-size: 20px;";
        btn2.onclick = function() {
            botDisableGB = true
            // hide both buttons
            btn2.style.display = "none";
            btn.style.display = "none";
        }
        document.body.appendChild(btn2);
    }
}