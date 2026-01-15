/**
 * Script de prueba para SharePoint 2019
 * Imprime un saludo y datos del contexto en la consola
 */
(function() {
    // Función principal
    function inicializarHolaMundo() {
        console.log("%c--- SCRIPT GLOBAL CARGADO ---", "color: #0078d4; font-weight: bold; font-size: 14px;");
        console.log("Hola Mundo desde SharePoint 2019");
        
        // Información de contexto útil para depuración
        if (typeof _spPageContextInfo !== "undefined") {
            console.log("Sitio actual: " + _spPageContextInfo.webAbsoluteUrl);
            console.log("Usuario: " + _spPageContextInfo.userDisplayName);
            console.log("Página: " + _spPageContextInfo.serverRequestPath);
        } else {
            console.warn("Aviso: _spPageContextInfo no está disponible aún.");
        }
        
        console.log("---------------------------------");
    }

    // Ejecutar cuando el DOM esté listo
    if (document.readyState === "complete" || document.readyState === "interactive") {
        inicializarHolaMundo();
    } else {
        document.addEventListener("DOMContentLoaded", inicializarHolaMundo);
    }
})();
