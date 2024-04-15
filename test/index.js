

const convertToXML = () => {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files ? fileInput.files[0] : null;
    const fileName = file.name.split('.').slice(0, -1).join('.');
    if (!file) {
        alert('Please select a file to convert.');
        return;
    }

    const formData = new FormData();
    formData.append('file', file);

    fetch('http://localhost:3001/', {
        method: 'POST',
        body: formData,
    })
        .then(response => response.blob())
        .then(blob => {
            const url = window.URL.createObjectURL(blob);
            // Create a new anchor element
            const a = document.createElement('a');
            a.href = url;
            a.download = `${fileName}.xml`;
            document.body.appendChild(a);
            a.click();
            a.remove();  // Remove the element after download
            window.URL.revokeObjectURL(url);  // Free up storage used by the blob
        })
        .catch(error => console.error('Error:', error));





    // alert('Conversion functionality is not implemented yet.');
}