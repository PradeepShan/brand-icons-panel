// Function to print messages to our HTML screen
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
    await PowerPoint.run(async (context) => {
        
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items");
        await context.sync();

        if (selectedShapes.items.length === 0) {
            logMessage("ABORT: No shape selected on the slide!");
            return;
        }

        logMessage("Shape found. Reading dimensions...");
        const oldIcon = selectedShapes.items[0];
        oldIcon.load(["top", "left", "width", "height", "rotation"]);
        await context.sync();

        const pTop = oldIcon.top;
        const pLeft = oldIcon.left;
        const pWidth = oldIcon.width;
        const pHeight = oldIcon.height;
        const pRotation = oldIcon.rotation;

        logMessage("Injecting new icon...");
        const currentSlide = context.presentation.getSelectedSlides().getItemAt(0);
        
        // This is usually where PowerPoint gets picky about SVGs
        const newIcon = currentSlide.shapes.addImage(base64Image);

        newIcon.top = pTop;
        newIcon.left = pLeft;
        newIcon.width = pWidth;
        newIcon.height = pHeight;
        newIcon.rotation = pRotation;

        logMessage("Deleting old icon...");
        oldIcon.delete();

        await context.sync();
        logMessage("SUCCESS: Swap complete!");
        
    }).catch(function (error) {
        logMessage("API ERROR: " + error);
    });
}
