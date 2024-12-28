#target illustrator

function createPreviewMenu() {
    var dlg = new Window("dialog", "Crear Previo");
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    // Campo de texto para Cliente
    dlg.add("statictext", undefined, "Cliente:");
    var clienteInput = dlg.add("edittext", undefined, "", {multiline: false});
    clienteInput.preferredSize = [200, 23];

    // Casilla para activar/desactivar Material
    var materialCheckboxGroup = dlg.add("group");
    materialCheckboxGroup.orientation = "column";
    materialCheckboxGroup.add("statictext", undefined, "Material:");
    var materialCheckbox = materialCheckboxGroup.add("checkbox", undefined, "Incluir Material");
    materialCheckbox.value = true; // Por defecto está activada

    // Campo de Material
    var materialDropdown = dlg.add("dropdownlist", undefined, ["Dtf textil", "Dtf UV", "Lona brillante", "Lona mate", "Lona mesh", "Lona translúcida", "Tela canva", "Lona personalizada", "Vinil brillante", "Vinil mate", "Vinil transparente", "Vinil microperforado", "Vinil personalizado", "Corte de vinil brillante", "Corte de vinil mate", "Corte de vinil transparente", "Corte de vinil microperforado", "Corte de vinil personalizado", "Couche 130 gr Brillante", "Couche 130 gr Mate", "Couche 300 gr Brillante", "Couche 300 gr Mate", "Opalina 120 gr", "Opalina 216 gr", "Bond 75gr", "Papel Adhesivo Brillante", "Papel Adhesivo Mate", "Papel Sulfatado", "Otro..."]);
    materialDropdown.selection = 0;
    materialDropdown.enabled = materialCheckbox.value;

    // Campo para "Otro..." en Material
    var materialOtroInput = dlg.add("edittext", undefined, "", {multiline: false});
    materialOtroInput.preferredSize = [200, 23];
    materialOtroInput.visible = false;

    // Casilla para activar/desactivar Acabados
    var acabadosCheckboxGroup = dlg.add("group");
    acabadosCheckboxGroup.orientation = "column";
    acabadosCheckboxGroup.add("statictext", undefined, "Acabados:");
    var acabadosCheckbox = acabadosCheckboxGroup.add("checkbox", undefined, "Incluir Acabados");
    acabadosCheckbox.value = true; // Por defecto está activada

    // Campo de Acabados
    var acabadosDropdown = dlg.add("dropdownlist", undefined, ["Sin Acabados", "Solo dobladillo", "Dobladillo y ojillos", "Roll-Up", "Banner", "Tensar", "Refilado al ras", "Refilado por pieza", "Laminado sobre pvc", "Laminado sobre estireno", "Laminado sobre acrílico", "Laminado sobre MDF", "Sólo depilado", "Depilado y transfer", "Corte de router", "Corte de láser", "Acabado personalizado", "Otro..."]);
    acabadosDropdown.selection = 0;
    acabadosDropdown.enabled = acabadosCheckbox.value;

    // Campo para "Otro..." en Acabados
    var acabadosOtroInput = dlg.add("edittext", undefined, "", {multiline: false});
    acabadosOtroInput.preferredSize = [200, 23];
    acabadosOtroInput.visible = false;

    // Casilla para activar/desactivar Medidas
    var medidasCheckboxGroup = dlg.add("group");
    medidasCheckboxGroup.orientation = "column";
    medidasCheckboxGroup.add("statictext", undefined, "Dimensiones:");
    var medidasCheckbox = medidasCheckboxGroup.add("checkbox", undefined, "Incluir Medidas");
    medidasCheckbox.value = true; // Por defecto está activada

    // Campo de relleno automático de las medidas
    var dimensionesInput = dlg.add("edittext", undefined, "Calculando...", {multiline: false, readonly: true});
    dimensionesInput.preferredSize = [200, 30];
    dimensionesInput.enabled = medidasCheckbox.value;

	// Casilla para activar/desactivar Cantidades
    var CantidadesCheckboxGroup = dlg.add("group");
    CantidadesCheckboxGroup.orientation = "column";
    CantidadesCheckboxGroup.add("statictext", undefined, "Cantidades:");
    var CantidadesCheckboxGroup = CantidadesCheckboxGroup.add("checkbox", undefined, "Incluir Cantidades");
    CantidadesCheckboxGroup.value = true; // Por defecto está activada

    // Campo de texto para Cantidades
    dlg.add("statictext", undefined, "Cantidades:");
    var CantidadesInput = dlg.add("edittext", undefined, "", {multiline: false});
    CantidadesInput.preferredSize = [200, 23];

    // Campo de carpeta destino
    dlg.add("statictext", undefined, "Carpeta de destino de trabajos:");
    var carpetaDestinoButton = dlg.add("button", undefined, "Seleccionar Carpeta");
    var carpetaDestinoPath = dlg.add("edittext", undefined, "", {readonly: true});
    carpetaDestinoPath.preferredSize = [300, 30];

    carpetaDestinoButton.onClick = function () {
        var carpetaSeleccionada = Folder.selectDialog("Selecciona la carpeta de destino");
        if (carpetaSeleccionada) {
            carpetaDestinoPath.text = carpetaSeleccionada.fsName;
        }
    };

    // Espacio para botones
    var buttonGroup = dlg.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "right";

    var createButton = buttonGroup.add("button", undefined, "Crear Previo");
    var cancelButton = buttonGroup.add("button", undefined, "Cancelar");

    // Texto centrado en la parte inferior
    var footerGroup = dlg.add("group");
    footerGroup.orientation = "column";
    footerGroup.alignment = "center";
    footerGroup.add("statictext", undefined, "Desarrollo: Pedro M. /// Test para MQ Print México");

    // Obtener dimensiones de la mesa inicial
    var doc = app.activeDocument;
    if (doc) {
        var artboard = doc.artboards[0]; // Primer mesa de trabajo
        var rect = artboard.artboardRect; // Rectángulo de la mesa de trabajo

        // Cálculo de ancho y altura
        var width = rect[2] - rect[0]; // ancho (right - left)
        var height = rect[1] - rect[3]; // altura (top - bottom)

        // Convertir las medidas a cm si las unidades son en puntos (1 pulgada = 2.54 cm, 72 puntos por pulgada)
        var widthInCm = width * 2.54 / 72;
        var heightInCm = height * 2.54 / 72;

        dimensionesInput.text = Math.round(widthInCm) + "x" + Math.round(heightInCm) + "cm";
    } else {
        dimensionesInput.text = "No hay documento abierto";
    }

    // Mostrar campo "Otro..." para Material
    materialDropdown.onChange = function() {
        if (materialDropdown.selection.text === "Otro...") {
            materialOtroInput.visible = true;
        } else {
            materialOtroInput.visible = false;
        }
    };

    // Mostrar campo "Otro..." para Acabados
    acabadosDropdown.onChange = function() {
        if (acabadosDropdown.selection.text === "Otro...") {
            acabadosOtroInput.visible = true;
        } else {
            acabadosOtroInput.visible = false;
        }
    };

    // Eventos de botones
    createButton.onClick = function () {
        var cliente = clienteInput.text;
        var material = materialDropdown.selection.text === "Otro..." ? materialOtroInput.text : materialDropdown.selection.text;
        var acabados = acabadosDropdown.selection.text === "Otro..." ? acabadosOtroInput.text : acabadosDropdown.selection.text;
        var Cantidades = CantidadesInput.text;
        var dimensiones = dimensionesInput.text;
        var carpetaDestino = carpetaDestinoPath.text;

        if (!cliente || !material || !acabados || !dimensiones || !Cantidades || !carpetaDestino) {
            alert("Por favor, rellena todos los campos.");
            return;
        }
		
		// Generar el nombre de la carpeta basado en los campos habilitados
        var carpetaFinalNombre = cliente;
        if (materialCheckbox.value && material) {
            carpetaFinalNombre += " " + material;
        }
        if (acabadosCheckbox.value && acabados) {
            carpetaFinalNombre += " " + acabados;
        }
        if (medidasCheckbox.value) {
            carpetaFinalNombre += " " + dimensiones;
        }
        if (Cantidades) {
            carpetaFinalNombre += " " + Cantidades;
        }

        // Obtener fecha actual
        var fecha = new Date();
        var meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
        var mesActual = meses[fecha.getMonth()];
        var dia = fecha.getDate();

        // Crear estructura de carpetas
        var carpetaCliente = new Folder(carpetaDestino + "/" + cliente);
        if (!carpetaCliente.exists) carpetaCliente.create();

        var carpetaMes = new Folder(carpetaCliente.fsName + "/" + mesActual);
        if (!carpetaMes.exists) carpetaMes.create();

        var carpetaDia = new Folder(carpetaMes.fsName + "/" + dia);
        if (!carpetaDia.exists) carpetaDia.create();

        var carpetaFinalNombre = material + " " + acabados + " " + dimensiones + " " + Cantidades + " " + cliente;
        var carpetaFinal = new Folder(carpetaDia.fsName + "/" + carpetaFinalNombre);
        if (!carpetaFinal.exists) {
            if (carpetaFinal.create()) {
                alert("Estructura de carpetas creada con éxito!" + "\n" + "Revisa que tu carpeta de salida contenga el cliente, el mes, el día y el proyecto.:\n" + carpetaFinal.fsName);

				// Abrir la carpeta en el explorador de archivos
				carpetaFinal.execute();
	
                // Guardar el archivo actual en formato .ai , .tif y .pdf
                if (doc) {
                    var archivoBaseNombre = carpetaFinalNombre;
                    var pdfFile = new File(carpetaFinal.fsName + "/" + archivoBaseNombre + ".pdf");
                    var tifFile = new File(carpetaFinal.fsName + "/" + archivoBaseNombre + ".tif");
                    var aiFile = new File(carpetaFinal.fsName + "/" + archivoBaseNombre + ".ai");

                    // Guardar en formato .pdf
                    var pdfSaveOptions = new PDFSaveOptions();
                    pdfSaveOptions.compatibility = PDFCompatibility.ACROBAT6; // predeterminado con Acrobat 6 (PDF 1.5) o el preajuste seleccionado en la config del usuario
                    pdfSaveOptions.preserveEditability = false; // Evitar que se guarde con capacidad de edición (pa los malayas rekls)
                    pdfSaveOptions.optimization = true; // Compresión LZW predeterminada pa reducir tamaño del archivo
                    doc.saveAs(pdfFile, pdfSaveOptions);

                    // Exportar en formato .tif
                    var tifExportOptions = new ExportOptionsTIFF();
                    tifExportOptions.resolution = 200;
                    tifExportOptions.saveMultipleArtboards = true;
                    tifExportOptions.byteOrder = TIFFByteOrder.IBMPC;
                    doc.exportFile(tifFile, ExportType.TIFF, tifExportOptions);

                    // Guardar en formato .ai
                    var aiSaveOptions = new IllustratorSaveOptions();
                    doc.saveAs(aiFile, aiSaveOptions);

                    alert("Archivos guardados:\n" + aiFile.fsName + "\n" + tifFile.fsName + "\n" + pdfFile.fsName);
                }
            } else {
                alert("Error al crear la estructura de carpetas. Verifica la ruta de destino.");
            }
        } else {
            alert("La carpeta final ya existe:\n" + carpetaFinal.fsName);
        }

        dlg.close();
    };

    cancelButton.onClick = function () {
        dlg.close();
    };

    dlg.show();
}

createPreviewMenu();
