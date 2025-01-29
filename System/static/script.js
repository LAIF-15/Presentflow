// Function to load HTML content dynamically and insert into the DOM
async function loadContent() {
    try {
        // Fetch content for each section using Flask routes
        const topBar = await fetch('/topbar.html').then(res => res.text());
        const sidebar = await fetch('/sidebar.html').then(res => res.text());
        const mainContent = await fetch('/maincontent.html').then(res => res.text());
        const aboutSystem = await fetch('/aboutsystem.html').then(res => res.text());
        const aboutUs = await fetch('/aboutus.html').then(res => res.text());

        // Insert fetched content into respective elements
        document.getElementById('top-bar').innerHTML = topBar;
        document.getElementById('sidebar').innerHTML = sidebar;
        document.getElementById('main-content').innerHTML = mainContent;
        document.getElementById('about-system').innerHTML = aboutSystem;
        document.getElementById('about-us').innerHTML = aboutUs;
    } catch (error) {
        console.error("Error loading content:", error);
    }
}

// Call the function to load content when the page loads
window.onload = async function () {
    await loadContent(); // Ensure content is loaded first

    const fileDropArea = document.getElementById('file-drop-area');
    const fileInput = document.getElementById('file-upload');
    const uploadButton = document.getElementById('upload-btn');
    const fileNameDisplay = document.getElementById('file-name-display'); // Add this line to select the file name display element

    // Enable upload button and display file name when a file is selected
    fileInput.addEventListener('change', function() {
        if (fileInput.files.length > 0) {
            const fileName = fileInput.files[0].name;
            console.log("File selected:", fileName);

            // Update the file name display
            fileNameDisplay.textContent = `${fileName}`;

            uploadButton.disabled = false; // Enable the upload button
        } else {
            fileNameDisplay.textContent = 'Drag & Drop your file here or click to select'; // Clear display if no file is selected
            uploadButton.disabled = true; // Disable button if no file
        }
    });

    // Drag & Drop functionality
    fileDropArea.addEventListener('dragover', function(event) {
        event.preventDefault(); // Allow dropping
        fileDropArea.classList.add('drag-over'); // Add drag-over class for styling
    });

    fileDropArea.addEventListener('dragleave', function() {
        fileDropArea.classList.remove('drag-over'); // Remove drag-over class when dragging leaves
    });

    fileDropArea.addEventListener('drop', function(event) {
        event.preventDefault(); // Prevent default behavior (open file)
        fileDropArea.classList.remove('drag-over'); // Remove drag-over class

        const files = event.dataTransfer.files;
        if (files.length > 0) {
            fileInput.files = files; // Assign dropped files to the file input

            const fileName = files[0].name;
            console.log("File dropped:", fileName);

            // Update the file name display
            fileNameDisplay.textContent = `${fileName}`;

            uploadButton.disabled = false; // Enable the upload button
        }
    });

    // Function to handle file upload
    function uploadFile() {
        const file = fileInput.files[0];
        if (file) {
            console.log(`Uploading: ${file.name}`);
            alert(`Uploading: ${file.name}`);
            // Implement file upload logic here (e.g., send the file to the server using FormData)
        } else {
            alert("No file selected for upload.");
        }
    }
    
    const modal = document.getElementById('instructions-modal');
    const closeModalButton = document.getElementById('close-modal-btn'); // Optional: close button for modal
    const instructionButton = document.getElementById('instruction-button'); // The button to open modal

    // Close the modal when the close button is clicked (optional)
    if (closeModalButton) {
        closeModalButton.addEventListener('click', () => {
            modal.style.display = 'none'; // Hide modal
        });
    }

    // Show the modal when the "How to Use" button is clicked
    instructionButton.addEventListener('click', () => {
        modal.style.display = 'flex'; // Display modal (flex for center alignment)
    });

    // Optional: Close the modal if the user clicks outside of the modal content
    modal.addEventListener('click', (e) => {
        if (e.target === modal) {
            modal.style.display = 'none'; // Hide modal if clicked outside the content
        }
    });

    const modal_download = document.getElementById('download-modal');
    const closeModalButton1 = document.getElementById('close-modal-btn1'); 
    const downloadButton = document.getElementById('download-button');

    if (closeModalButton1) {
        closeModalButton1.addEventListener('click', () => {
            modal_download.style.display = 'none'; 
        });
    }

    if (downloadButton) {
        downloadButton.addEventListener('click', () => {
            modal_download.style.display = 'flex'; 
        });
    }
    
    if (modal_download) {
        modal_download.addEventListener('click', (e) => {
            if (e.target === modal_download) {
                modal_download.style.display = 'none';
            }
        });
    }
};

document.addEventListener('DOMContentLoaded', () => {
    const mainPreview = document.getElementById('main-preview');
    const currentSlideText = document.getElementById('current-slide');
    const thumbnailItems = document.querySelectorAll('.thumbnail-item');

    thumbnailItems.forEach((item) => {
        item.addEventListener('click', () => {
            // Get the slide data
            const base64Image = item.getAttribute('data-image');
            const slideIndex = item.getAttribute('data-slide');

            mainPreview.src = base64Image;

            currentSlideText.textContent = slideIndex;

            thumbnailItems.forEach((thumb) => thumb.classList.remove('active'));
            item.classList.add('active');
        });
    });

    if (thumbnailItems.length > 0) {
        thumbnailItems[0].classList.add('active');
    }
});
