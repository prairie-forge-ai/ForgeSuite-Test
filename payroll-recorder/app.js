(function () {
    const bundleSrc = 'app.bundle.js?v=miov6em0';
    let bootStarted = false;

    function loadBundle() {
        if (bootStarted) return;
        bootStarted = true;
        const script = document.createElement('script');
        script.src = bundleSrc;
        script.async = false;
        script.onerror = (error) => displayError('Unable to load Payroll Recorder. Please refresh.', error);
        document.body.appendChild(script);
    }

    function displayError(message, error) {
        console.error(message, error);
        const loading = document.getElementById('loading') || createFallbackContainer();
        loading.textContent = message;
        loading.style.color = '#f87171';
    }

    function createFallbackContainer() {
        const div = document.createElement('div');
        div.id = 'loading';
        div.style.fontFamily = "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif";
        div.style.padding = '20px';
        div.style.textAlign = 'center';
        document.body.appendChild(div);
        return div;
    }

    function bootWhenReady() {
        if (typeof Office !== 'undefined' && typeof Office.onReady === 'function') {
            Office.onReady()
                .then(loadBundle)
                .catch((error) => {
                    console.warn('Office.onReady failed, falling back to immediate boot.', error);
                    loadBundle();
                });
            return;
        }

        if (document.readyState === 'complete' || document.readyState === 'interactive') {
            loadBundle();
        } else {
            document.addEventListener('DOMContentLoaded', loadBundle);
        }
    }

    bootWhenReady();
})();
