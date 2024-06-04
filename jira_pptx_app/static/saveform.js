document.addEventListener('DOMContentLoaded', (event) => {
  const downloadButton = document.getElementById('downloadButton');
  const uploadButton = document.getElementById('uploadButton');
  const fileInput = document.getElementById('fileInput');
  fileInput.type = 'file';
  fileInput.style.display = 'none';
  document.body.appendChild(fileInput);

  if (downloadButton) {
    downloadButton.addEventListener('click', () => {
      console.log('Download function called');
      const data = {
        jira_server_url: document.getElementById('jira_server_url').value,
        email: document.getElementById('email').value,
        // api_key: document.getElementById('api_key').value,
        jql_query: document.getElementById('jql_query').value,
        fields_to_include: document.getElementById('fields_to_include').value,
        sort_field: document.getElementById('sort_field').value,
        sort_order: document.getElementById('sort_order').value,
        fields_to_display: document.getElementById('fields_to_display').value,
        fix_version_ce: document.getElementById('fix_version_ce').value,
        fix_version_ee: document.getElementById('fix_version_ee').value,
        rows_per_slide: document.getElementById('rows_per_slide').value,
        // Add any additional form fields here
      };

      const dataBlob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
      const dataUrl = URL.createObjectURL(dataBlob);

      const a = document.createElement('a');
      a.href = dataUrl;
      a.download = 'formData.json';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);

      setTimeout(() => {
        URL.revokeObjectURL(dataUrl);
      }, 100);
    });
  } else {
    console.error('Download button not found');
  }

  if (uploadButton) {
    uploadButton.addEventListener('click', () => {
      console.log('Upload function called');
      fileInput.click();
    });

    fileInput.addEventListener('change', (event) => {
      console.log('File input changed');
      const file = event.target.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (e) => {
          const text = e.target.result;
          try {
            const data = JSON.parse(text);
            console.log(data);

            document.getElementById('jira_server_url').value = data.jira_server_url || '';
            document.getElementById('email').value = data.email || '';
            document.getElementById('api_key').value = data.api_key || '';
            document.getElementById('jql_query').value = data.jql_query || '';
            document.getElementById('fields_to_include').value = data.fields_to_include || '';
            document.getElementById('sort_field').value = data.sort_field || '';
            document.getElementById('sort_order').value = data.sort_order || '';
            document.getElementById('fields_to_display').value = data.fields_to_display || '';
            document.getElementById('fix_version_ce').value = data.fix_version_ce || '';
            document.getElementById('fix_version_ee').value = data.fix_version_ee || '';
            document.getElementById('rows_per_slide').value = data.rows_per_slide || '';
            // Add any additional form fields here

          } catch (err) {
            console.error('Error parsing JSON:', err);
          }
        };
        reader.readAsText(file);
      }
    });
  } else {
    console.error('Upload button not found');
  }
});
