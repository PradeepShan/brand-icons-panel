// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("PowerPoint Add-in Ready");
    }
});

// Step 1: Fetch the SVG and convert it to Base64
async function processIconClick(svgUrl) {
    try {
        const response = await fetch(svgUrl);
        const blob = await response.blob();
        const reader = new FileReader();
        
        reader.onloadend = async function() {
            // PowerPoint requires the raw base64 string without the prefix
            const base64String = reader.result.split(',')[1];
            await executeSwap(base64String);
        };
        reader.readAsDataURL(blob);
    } catch (error) {
        console.error("Error loading SVG: ", error);
    }
}

// Step 2: Talk to PowerPoint to swap the shapes
async function executeSwap(base64Image) {
    await PowerPoint.run(async (context) => {
        // Find the selected icon
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items");
        await context.sync();

        if (selectedShapes.items.length === 0) {
            console.error("Please select an icon on the slide first.");
            return;
        }

        const oldIcon = selectedShapes.items[0];
        
        // Read exact dimensions
        oldIcon.load(["top", "left", "width", "height", "rotation"]);
        await context.sync();

        const pTop = oldIcon.top;
        const pLeft = oldIcon.left;
        const pWidth = oldIcon.width;
        const pHeight = oldIcon.height;
        const pRotation = oldIcon.rotation;

        // Target current slide
        const currentSlide = context.presentation.getSelectedSlides().getItemAt(0);

        // Inject new image
        const newIcon = currentSlide.shapes.addImage(base64Image);

        // Apply DNA
        newIcon.top = pTop;
        newIcon.left = pLeft;
        newIcon.width = pWidth;
        newIcon.height = pHeight;
        newIcon.rotation = pRotation;

        // Delete old icon
        oldIcon.delete();

        await context.sync();
    });
}