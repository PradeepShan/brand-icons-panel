Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        // Ready!
    }
});

async function processIconClick(svgUrl) {
    try {
        const response = await fetch(svgUrl);
        if (!response.ok) {
            alert("Error: GitHub couldn't find the SVG at " + svgUrl);
            return;
        }
        const blob = await response.blob();
        const reader = new FileReader();
        
        reader.onloadend = async function() {
            try {
                const base64String = reader.result.split(',')[1];
                await executeSwap(base64String);
            } catch (err) {
                alert("Swap Engine Error: " + err.message);
            }
        };
        reader.readAsDataURL(blob);
    } catch (error) {
        alert("Fetch Error: " + error.message);
    }
}

async function executeSwap(base64Image) {
    await PowerPoint.run(async (context) => {
        // 1. Check for selected shape
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items");
        await context.sync();

        if (selectedShapes.items.length === 0) {
            alert("Please select the old icon on your slide first!");
            return;
        }

        const oldIcon = selectedShapes.items[0];
        
        // 2. Read dimensions
        oldIcon.load(["top", "left", "width", "height", "rotation"]);
        await context.sync();

        const pTop = oldIcon.top;
        const pLeft = oldIcon.left;
        const pWidth = oldIcon.width;
        const pHeight = oldIcon.height;
        const pRotation = oldIcon.rotation;

        // 3. Target slide and swap
        const currentSlide = context.presentation.getSelectedSlides().getItemAt(0);
        const newIcon = currentSlide.shapes.addImage(base64Image);

        newIcon.top = pTop;
        newIcon.left = pLeft;
        newIcon.width = pWidth;
        newIcon.height = pHeight;
        newIcon.rotation = pRotation;

        oldIcon.delete();
        await context.sync();
        
    }).catch(function (error) {
        // This catches API errors (like PowerPoint refusing the SVG format)
        alert("PowerPoint API Error: " + error);
    });
}
