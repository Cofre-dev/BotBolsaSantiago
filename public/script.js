const fechaInput = document.getElementById('fecha');
const btnDescargar = document.getElementById('btn-descargar');
const errorMsg = document.getElementById('error-msg');
const statusContainer = document.getElementById('status-container');
const progressFill = document.getElementById('progress-fill');
const statusText = document.getElementById('status-text');
const logs = document.getElementById('logs');
const btnLoader = document.getElementById('btn-loader');

// Set max date to today
const today = new Date().toISOString().split('T')[0];
fechaInput.setAttribute('max', today);

btnDescargar.addEventListener('click', async () => {
    const selectedDate = new Date(fechaInput.value);

    // Validate if it's a weekday
    const day = selectedDate.getUTCDay(); // 0 = Sunday, 6 = Saturday

    if (!fechaInput.value) {
        errorMsg.textContent = 'Por favor selecciona una fecha.';
        return;
    }

    if (day === 0 || day === 6) {
        errorMsg.textContent = 'Por favor selecciona un día hábil (Lunes a Viernes).';
        return;
    }

    errorMsg.textContent = '';
    startProcess(fechaInput.value);
});

async function startProcess(date) {
    btnDescargar.disabled = true;
    btnLoader.style.display = 'block';
    statusContainer.classList.remove('hidden');
    progressFill.style.width = '10%';
    statusText.textContent = 'Iniciando navegación...';

    try {
        const response = await fetch('/api/process', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ date })
        });

        const reader = response.body.getReader();
        const decoder = new TextDecoder();

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            const chunk = decoder.decode(value);
            const lines = chunk.split('\n');

            for (const line of lines) {
                if (!line) continue;
                try {
                    const data = JSON.parse(line);
                    if (data.status) statusText.textContent = data.status;
                    if (data.progress) progressFill.style.width = data.progress + '%';

                    if (data.fileUrl) {
                        progressFill.style.width = '100%';
                        statusText.textContent = '¡Descarga completada!';

                        // Create download link
                        const a = document.createElement('a');
                        a.href = data.fileUrl;
                        a.download = data.fileName || 'boletin_bursatil.xlsx';
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                    }

                    if (data.error) {
                        statusText.textContent = 'Error: ' + data.error;
                        progressFill.style.background = '#d63031';
                    }
                } catch (e) {
                    console.log(e);
                }
            }
        }
    } catch (err) {
        statusText.textContent = 'Fallo de conexión con el servidor.';
        console.error(err);
    } finally {
        btnDescargar.disabled = false;
        btnLoader.style.display = 'none';
    }
}
