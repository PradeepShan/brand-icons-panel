function logMessage(msg) {
    const screen = document.getElementById('debugScreen');
    if (screen) {
        screen.innerHTML += "<br>> " + msg;
    }
}

// 1. Run this immediately when PowerPoint connects
Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        logMessage("PowerPoint connected successfully.");
        loadIconsFromGitHub();
    }
});

// 2. THE NEW AUTO-LOADER ENGINE
async function loadIconsFromGitHub() {
    // CHANGE THESE TWO VARIABLES TO MATCH YOUR GITHUB INFO
    const githubUser = "YOUR_GITHUB_USERNAME";
    const githubRepo = "YOUR_REPO_NAME"; 
    
    const apiUrl = `https://api.github.com/repos/${githubUser}/${githubRepo}/contents/assets`;
    logMessage("Scanning folder: " + apiUrl);

    try {
        const response = await fetch(apiUrl);
        if (!response.ok) throw new Error("Could not read folder. Is the repo public?");
        
        const files = await response.json();
        const grid = document.getElementById('iconGrid');
        grid.innerHTML = ""; // Clear the "Loading..." text
        
        // Loop through the folder and build the HTML for each SVG
        files.forEach(file => {
            if (file.name.endsWith('.svg') || file.name.endsWith('.png')) {
                // file.download_url gives us the direct link to the raw image file
                const html = `
                    <div class="icon-card" onclick="processIconClick('${file.download_url}')">
                        <img src="${file.download_url}" alt="${file.name}">
                        <div class="icon-name">${file.name.replace('.svg', '')}</div>
                    </div>
                `;
                grid.innerHTML += html;
            }
        });
        
        logMessage(`Successfully loaded ${files.length} files.`);
        
    } catch (error) {
        logMessage("Auto-load Error: " + error.message);
    }
}

// 3. THE SWAP ENGINE (Unchanged, just uses the new URL)
async function processIconClick(svgUrl) {
    logMessage("Fetching: " + svgUrl);
    try {
        const response = await fetch(svgUrl);
        if (!response.ok) {
            logMessage("ERROR: Couldn't download image.");
            return;
        }
        
        const blob = await response.blob();
        const reader = new FileReader();
        
        reader.onloadend = async function() {
            try {
                const base64String = reader.result.split(',')[1];
                await executeSwap(base64String);
            } catch (err) {
                logMessage("Engine Error: " + err.message);
            }
        };
        reader.readAsDataURL(blob);
    } catch (error) {
        logMessage("Fetch Error: " + error.message);
    }
}

async function executeSwap(base64Image) {
    logMessage("Starting slide swap process...");
    let pTop, pLeft, pWidth, pHeight;

    try {
        await PowerPoint.run(async (context) => {
            const selectedShapes = context.presentation.getSelectedShapes();
            selectedShapes.load("items");
            await context.sync();

            if (selectedShapes.items.length === 0) {
                throw new Error("No shape selected on the slide!");
            }

            const oldIcon = selectedShapes.items[0];
            oldIcon.load(["top", "left", "width", "height"]);
            await context.sync();

            pTop = oldIcon.top;
            pLeft = oldIcon.left;
            pWidth = oldIcon.width;
            pHeight = oldIcon.height;

            oldIcon.delete();
            await context.sync();
        });
        
        Office.context.document.setSelectedDataAsync(
            base64Image,
            { coercionType: Office.CoercionType.Image, imageLeft: pLeft, imageTop: pTop, imageWidth: pWidth, imageHeight: pHeight },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    logMessage("API ERROR: " + asyncResult.error.message);
                } else {
                    logMessage("SUCCESS: Swap complete!");
                }
            }
        );
    } catch (error) {
        logMessage("ERROR: " + error.message);
    }
}
