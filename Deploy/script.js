(function() {
    let isProcessing = false;

    function getCorrectLibraryUrl() {
        try {
            // Intentar obtener el ID del WebPart con foco
            let wpqId = (typeof g_currentWPQ !== 'undefined' && g_currentWPQ) ? g_currentWPQ : null;
            
            if (wpqId) {
                let numericId = wpqId.replace('WPQ', '');
                let currentCtx = window['ctx' + numericId];
                if (currentCtx && currentCtx.listUrlDir) {
                    return window.location.origin + currentCtx.listUrlDir;
                }
            }
            // Fallback al contexto general (ctx suele ser el primero de la página)
            if (typeof ctx !== 'undefined' && ctx.listUrlDir) {
                return window.location.origin + ctx.listUrlDir;
            }
        } catch (e) { console.error("Error obteniendo URL:", e); }
        return null;
    }

    function forceEnableButton() {
        if (isProcessing) return;
        isProcessing = true;

        const btn = document.getElementById('Ribbon.Library.Actions.OpenWithExplorer-Medium');
        if (btn) {
            // Si el botón tiene la clase de deshabilitado, se la quitamos a él y al padre (li)
            if (btn.classList.contains('ms-cui-disabled')) {
                btn.classList.remove('ms-cui-disabled');
                if (btn.parentElement) btn.parentElement.classList.remove('ms-cui-disabled');
                btn.setAttribute('aria-disabled', 'false');
            }
            
            // En lugar de btn.onclick, usamos un atributo para identificarlo
            btn.setAttribute('data-explorer-patched', 'true');
        }
        
        setTimeout(() => { isProcessing = false; }, 100);
    }

    // DELEGACIÓN DE EVENTOS: Escuchamos en el documento el clic para adelantarnos a SharePoint
    document.addEventListener('click', function(e) {
        // Buscamos si el clic (o el burbujeo) viene del botón de explorador
        const btn = e.target.closest('#Ribbon.Library.Actions.OpenWithExplorer-Medium');
        
        if (btn) {
            e.preventDefault();
            e.stopImmediatePropagation(); // Detiene los scripts originales de SharePoint

            let rawUrl = getCorrectLibraryUrl();
            if (rawUrl) {
                let cleanUrl = rawUrl.split('/Forms/')[0];
                
                // Si por error es una lista, avisamos
                if (cleanUrl.toLowerCase().includes('/lists/')) {
                    alert("El Explorador de archivos solo funciona en Bibliotecas de Documentos.");
                } else {
                    console.log("Iniciando transferencia a: " + cleanUrl);
                    window.location.href = "viewinfileexplorer:" + cleanUrl;
                }
            }
            return false;
        }
    }, true); // El 'true' activa la fase de captura (antes que los eventos de SP)

    // Observador para mantener el botón visualmente habilitado
    const targetNode = document.getElementById('s4-ribbonrow') || document.body;
    const observer = new MutationObserver(() => forceEnableButton());

    observer.observe(targetNode, { 
        childList: true, 
        subtree: true, 
        attributes: true, 
        attributeFilter: ['class'] 
    });

    // Primera ejecución
    setTimeout(forceEnableButton, 1000);
})();
