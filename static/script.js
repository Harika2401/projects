document.getElementById('uploadBtn').addEventListener('click', function() {
    document.getElementById('uploadOptions').style.display = 'block';
});

document.getElementById('cameraBtn').addEventListener('click', function() {
    const form = document.createElement('form');
    form.action = '/upload';
    form.method = 'post';
    const input = document.createElement('input');
    input.type = 'hidden';
    input.name = 'camera';
    input.value = 'true';
    form.appendChild(input);
    document.body.appendChild(form);
    form.submit();
});
