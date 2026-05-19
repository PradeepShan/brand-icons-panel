function logMessage(msg) {
    const screen = document.getElementById('debugScreen');
    if (screen) {
        screen.innerHTML += "<br>> " + msg;
    }
}

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        logMessage("PowerPoint connected successfully.");
    }
});

async function processIconClick(svgUrl) {
    logMessage("Fetching: " + svgUrl);
    try {
        const response = await fetch(svgUrl);
        if (!response.ok) {
            logMessage("ERROR: GitHub couldn't find " + svgUrl);
            return;
        }
        
        const blob = await response.blob();
        const reader = new FileReader();
        
        reader.onloadend = async function() {
            try {
                logMessage("SVG downloaded. Converting to Base64...");
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
        // STEP 1: Read the old icon's DNA and delete it
        await PowerPoint.run(async (context) => {
            const selectedShapes = context.presentation.getSelectedShapes();
            selectedShapes.load("items");
            await context.sync();

            if (selectedShapes.items.length === 0) {
                throw new Error("No shape selected on the slide!");
            }

            logMessage("Shape found. Reading dimensions...");
            const oldIcon = selectedShapes.items[0];
            oldIcon.load(["top", "left", "width", "height"]);
            await context.sync();

            pTop = oldIcon.top;
            pLeft = oldIcon.left;
            pWidth = oldIcon.width;
            pHeight = oldIcon.height;

            logMessage("Deleting old icon...");
            oldIcon.delete();
            await context.sync();
        });

        // STEP 2: Use the bulletproof Common API to inject the new icon
        logMessage("Injecting new icon via Common API...");
        
        Office.context.document.setSelectedDataAsync(
            base64Image,
            {
                coercionType: Office.CoercionType.Image,
                imageLeft: pLeft,
                imageTop: pTop,
                imageWidth: pWidth,
                imageHeight: pHeight
            },
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
