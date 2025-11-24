const body = document.body;
const toggle = document.getElementById('themeToggle');
const icon = document.getElementById('themeIcon');

toggle.addEventListener('click', () => {
    const isLight = body.getAttribute('data-theme') === 'light';
    body.setAttribute('data-theme', isLight ? 'dark' : 'light');
    icon.className = isLight ? 'fas fa-sun' : 'fas fa-moon';
});

document.getElementById('submitBtn').addEventListener('click', async () => {
    const data = {
        result_url: document.getElementById('result_url').value.trim(),
        start_roll: document.getElementById('start_roll').value.trim(),
        end_roll: document.getElementById('end_roll').value.trim(),
        delay: parseFloat(document.getElementById('delay').value) || 3.5
    };

    const total = parseInt(data.end_roll) - parseInt(data.start_roll) + 1;
    if (!data.result_url || !data.start_roll || !data.end_roll) {
        showAlert('error', 'All fields are required!');
        return;
    }
    if (total > 100 || total < 1) {
        showAlert('error', 'Maximum 100 roll numbers allowed!');
        return;
    }

    if (!confirm(`COBRA TECH will scrape ${total} results\nEstimated time: ~${Math.round(total * data.delay / 60)} minutes\nProceed?`)) return;

    document.getElementById('submitBtn').disabled = true;
    document.getElementById('progressContainer').style.display = 'block';
    document.getElementById('success').style.display = document.getElementById('error').style.display = 'none';

    function showAlert(type, msg) {
        const el = document.getElementById(type);
        el.textContent = msg;
        el.style.display = 'block';
    }

    try {
        const res = await fetch('/scrape', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        if (!res.ok) throw new Error((await res.json()).error || 'Failed');

        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `CT_REBONIZER_${data.start_roll}_to_${data.end_roll}_${new Date().toISOString().slice(0,10)}.xlsx`;
        a.click();

        showAlert('success', `COBRA TECH SUCCESS! ${total} results exported!`);
    } catch (e) {
        showAlert('error', e.message || 'Engine failed');
    } finally {
        document.getElementById('submitBtn').disabled = false;
        setTimeout(() => document.getElementById('progressContainer').style.display = 'none', 4000);
    }
});