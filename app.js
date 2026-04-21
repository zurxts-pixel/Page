const STATE = {
  records: [],
  headers: [
    "Consecutivo", 
    "Estado", 
    "Destino", 
    "Marcas", 
    "Número", 
    "Faltante (Tons)", 
    "Fecha siniestro ", 
    " Confronta", // Column H (Fecha)
    "Fecha de",   // Column I (Document type/id)
    " Confronta", // Column J (Document id)
    "Remisión", 
    "Item", 
    "Confronta / Acta de Hechos"
  ]
};

let currentSearch = "";

// 1. Persistencia Temporal: Cargar desde LocalStorage
function initPersistedState() {
  try {
    const saved = localStorage.getItem("confrontas_data");
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed)) {
        STATE.records = parsed;
      }
    }
  } catch (error) {
    console.error("No se pudo parsear el localStorage", error);
    showToast("Error al restaurar respaldo local", "error");
  }
}

// Persistencia Temporal: Guardar al LocalStorage
function persistState() {
  try {
    localStorage.setItem("confrontas_data", JSON.stringify(STATE.records));
  } catch (error) {
    console.error("Error al guardar en localStorage", error);
    showToast("Advertencia: No se pudo guardar la copia de seguridad local", "warning");
  }
}

// Auto-increment logic
function getNextConsecutivo() {
  let nextValue = 9; 
  if (STATE.records.length > 0) {
    const max = Math.max(...STATE.records.map(r => Number(r.consecutivo) || 0));
    nextValue = max + 1;
  }
  return nextValue;
}

function dateToExcelCell(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return "";
  const parts = dateStr.split('-');
  if (parts.length === 3) {
    return new Date(parts[0], parts[1] - 1, parts[2]);
  }
  return dateStr;
}

// Inicialización de Dark Mode
function initTheme() {
  const themeIconLight = document.getElementById('theme-icon-light');
  const themeIconDark = document.getElementById('theme-icon-dark');
  const themeIconLightMobile = document.getElementById('theme-icon-light-mobile');
  const themeIconDarkMobile = document.getElementById('theme-icon-dark-mobile');
  
  function applyTheme(isDark) {
    if (isDark) {
      document.documentElement.classList.add('dark');
      if (themeIconLight) themeIconLight.classList.remove('hidden');
      if (themeIconDark) themeIconDark.classList.add('hidden');
      if (themeIconLightMobile) themeIconLightMobile.classList.remove('hidden');
      if (themeIconDarkMobile) themeIconDarkMobile.classList.add('hidden');
    } else {
      document.documentElement.classList.remove('dark');
      if (themeIconLight) themeIconLight.classList.add('hidden');
      if (themeIconDark) themeIconDark.classList.remove('hidden');
      if (themeIconLightMobile) themeIconLightMobile.classList.add('hidden');
      if (themeIconDarkMobile) themeIconDarkMobile.classList.remove('hidden');
    }
  }

  // Verificar LocalStorage o Preferencia del Sistema
  const savedTheme = localStorage.getItem('theme');
  let isDark = false;
  
  if (savedTheme) {
    isDark = savedTheme === 'dark';
  } else {
    isDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
  }
  
  applyTheme(isDark);

  const toggleHandler = () => {
    const currentlyDark = document.documentElement.classList.contains('dark');
    const newThemeDark = !currentlyDark;
    localStorage.setItem('theme', newThemeDark ? 'dark' : 'light');
    applyTheme(newThemeDark);
  };

  const desktopToggle = document.getElementById('theme-toggle');
  const mobileToggle = document.getElementById('theme-toggle-mobile');
  
  if (desktopToggle) desktopToggle.addEventListener('click', toggleHandler);
  if (mobileToggle) mobileToggle.addEventListener('click', toggleHandler);
}


// Inicialización de la Sidebar Colapsable
function initSidebar() {
  const sidebar = document.getElementById('sidebar');
  const btnToggle = document.getElementById('btn-toggle-sidebar');
  
  if (!sidebar || !btnToggle) return;

  function setSidebarState(isCollapsed) {
    if (isCollapsed) {
      document.body.classList.add('sidebar-collapsed');
    } else {
      document.body.classList.remove('sidebar-collapsed');
    }
    localStorage.setItem('confrontas_sidebar_collapsed', isCollapsed);
  }

  // Cargar estado inicial
  const savedState = localStorage.getItem('confrontas_sidebar_collapsed') === 'true';
  setSidebarState(savedState);

  btnToggle.addEventListener('click', () => {
    const isCurrentlyCollapsed = document.body.classList.contains('sidebar-collapsed');
    setSidebarState(!isCurrentlyCollapsed);
  });
}

document.addEventListener("DOMContentLoaded", () => {
  lucide.createIcons();
  const form = document.getElementById("data-form");
  const consecutivoInput = document.getElementById("consecutivo");

  initTheme();
  initSidebar();
  initPersistedState();
  
  // Respetar valor si el usuario estaba escribiendo uno manualmente antes de refrescar
  const savedManual = localStorage.getItem("confrontas_manual_consecutivo");
  if (savedManual) {
    consecutivoInput.value = savedManual;
  } else {
    consecutivoInput.value = getNextConsecutivo();
  }

  // Si el usuario edita a mano, guardamos en persistencia temporal
  consecutivoInput.addEventListener("input", (e) => {
    localStorage.setItem("confrontas_manual_consecutivo", e.target.value);
  });
  
  // Render de inicio si hay datos
  renderTable();
  renderSummary();

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    
    try {
      if (!form.checkValidity()) {
        form.reportValidity();
        return;
      }

      const fd = new FormData(form);
      const faltanteStr = fd.get("faltante");
      const numeroStr = fd.get("numero");
      
      const numeroVal = Number(numeroStr);
      if (isNaN(numeroVal) || numeroVal < 0) {
        throw new Error("El campo 'Número' debe ser numérico y positivo.");
      }

      const faltanteVal = Number(faltanteStr);
      if (isNaN(faltanteVal)) {
        throw new Error("El Faltante debe ser numérico.");
      }
      
      if (faltanteVal < 0 || faltanteVal > 100000) {
        throw new Error("El valor de Faltante es irreal. Revisa los datos (rango permitido: 0-100,000).");
      }

      const record = {
        // Enlace oculto del objeto JS pero ID visible consecutivo
        id_unico: Date.now().toString(),
        consecutivo: Number(fd.get("consecutivo")),
        estado: fd.get("estado").trim().toUpperCase(),
        destino: fd.get("destino").trim().toUpperCase(),
        marcas: fd.get("marcas").trim().toUpperCase(),
        numero: numeroVal,
        faltante: faltanteVal,
        fecha_siniestro: fd.get("fecha_siniestro"),
        confronta_1: fd.get("confronta_1"), 
        fecha_de: fd.get("fecha_de").trim(),
        confronta_2: fd.get("confronta_2").trim(),
        remision: fd.get("remision").trim().toUpperCase(),
        item: fd.get("item").trim(),
        material: fd.get("material").trim().toUpperCase(),
        confrontada: false
      };

      STATE.records.push(record);
      persistState();
      
      renderTable();
      renderSummary();
      
      form.reset();
      
      // Limpiamos el guardado manual
      localStorage.removeItem("confrontas_manual_consecutivo");
      consecutivoInput.value = getNextConsecutivo();
      
      showToast("Registro guardado con éxito", "success");
    } catch (error) {
      showToast(error.message, "error");
    }
  });

  // Exportar Funcionalidad
  document.getElementById("btn-export").addEventListener("click", exportExcel);

  // Importar Excel Funcionalidad
  const fileImport = document.getElementById("file-import");
  let pendingImportRecords = [];
  
  document.getElementById("btn-import").addEventListener("click", () => {
    fileImport.click();
  });

  fileImport.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
      try {
        const data = new Uint8Array(evt.target.result);
        // cellDates: true permite que SheetJS intente parsear las celdas directamente como JS Date objetos nativos
        const workbook = XLSX.read(data, {type: 'array', cellDates: true});
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Obtener arreglos
        const aoa = XLSX.utils.sheet_to_json(worksheet, {header: 1});
        
        if (aoa.length < 2) {
            throw new Error("El archivo parece estar vacío o no tiene registros de datos.");
        }

        const fileHeaders = aoa[0];
        // Validar estrictamente la presencia del formato que hemos definido usando las primeras y sextas cabeceras (Faltante)
        if (fileHeaders[0] !== STATE.headers[0] || fileHeaders[5] !== STATE.headers[5]) {
            throw new Error("Formato incompatible. Las cabeceras no coinciden con la plantilla (A-M y O-P).");
        }

        pendingImportRecords = [];
        for (let i = 1; i < aoa.length; i++) {
           const row = aoa[i];
           // Saltarse posibles basuras de Excel o líneas en blanco. Tomamos "Consecutivo" (Índice 0) como pivote para existir.
           if (!row || row.length === 0 || row[0] === null || row[0] === undefined || row[0] === "") continue; 

           const formatExcelDate = (cell) => {
               if(!cell) return "";
               if(cell instanceof Date) {
                   const offset = cell.getTimezoneOffset();
                   cell = new Date(cell.getTime() - (offset * 60 * 1000));
                   return cell.toISOString().split('T')[0];
               }
               return String(cell);
           }

           const record = {
             id_unico: Date.now().toString() + Math.random().toString(36).substr(2, 5),
             consecutivo: Number(row[0]) || 0,
             estado: String(row[1] || "").trim().toUpperCase(),
             destino: String(row[2] || "").trim().toUpperCase(),
             marcas: String(row[3] || "").trim().toUpperCase(),
             numero: Number(row[4]) || 0,
             faltante: Number(row[5]) || 0,
             fecha_siniestro: formatExcelDate(row[6]),
             confronta_1: formatExcelDate(row[7]),
             fecha_de: String(row[8] || "").trim(),
             confronta_2: String(row[9] || "").trim(),
             remision: String(row[10] || "").trim().toUpperCase(),
             item: String(row[11] || "").trim(),
             material: String(row[12] || "").trim().toUpperCase(),
             confrontada: (String(row[12] || "").trim().toUpperCase() === "HECHO" || String(row[12] || "").trim().toUpperCase() === "SÍ")
           };

           pendingImportRecords.push(record);
        }

        if (pendingImportRecords.length === 0) {
            throw new Error("No se encontraron filas con datos válidos para procesar.");
        }

        // Desplegar Modal y guardar propuesta en pending
        document.getElementById("import-modal").classList.remove("hidden");

      } catch (err) {
        showToast(err.message || "Error al procesar el archivo Excel", "error");
        fileImport.value = ""; 
      }
    };
    reader.readAsArrayBuffer(file);
  });

  document.getElementById("btn-cancel-import").addEventListener("click", () => {
    document.getElementById("import-modal").classList.add("hidden");
    fileImport.value = "";
    pendingImportRecords = [];
  });

  document.getElementById("btn-confirm-import").addEventListener("click", () => {
    const action = document.querySelector('input[name="import_action"]:checked').value;
    
    if (action === "overwrite") {
      STATE.records = [...pendingImportRecords];
    } else {
      STATE.records.push(...pendingImportRecords);
    }

    persistState();
    localStorage.removeItem("confrontas_manual_consecutivo");
    renderTable();
    renderSummary();
    consecutivoInput.value = getNextConsecutivo();
    
    document.getElementById("import-modal").classList.add("hidden");
    fileImport.value = "";
    pendingImportRecords = [];
    
    showToast(`Registros importados con éxito. Módulo SUMIF actualizado en tiempo real.`, "success");
  });

  // Buscador Funcionalidad
  document.getElementById("search-input").addEventListener("input", (e) => {
    currentSearch = e.target.value.toLowerCase().trim();
    renderTable();
    renderSummary();
  });

  // Limpiar Todo Funcionalidad
  const modal = document.getElementById("clear-modal");
  document.getElementById("btn-clear").addEventListener("click", () => {
    modal.classList.remove("hidden");
  });
  
  document.getElementById("btn-cancel-clear").addEventListener("click", () => {
    modal.classList.add("hidden");
  });

  document.getElementById("btn-confirm-clear").addEventListener("click", () => {
    STATE.records = [];
    persistState();
    localStorage.removeItem("confrontas_manual_consecutivo");
    renderTable();
    renderSummary();
    consecutivoInput.value = getNextConsecutivo();
    modal.classList.add("hidden");
    showToast("Hoja de trabajo limpiada exitosamente.", "success");
  });

  // =============================================
  // Captura Rápida (Bulk Paste) Funcionalidad
  // =============================================
  const bulkModal = document.getElementById("bulk-modal");
  const bulkTextarea = document.getElementById("bulk-textarea");
  const bulkStepInput = document.getElementById("bulk-step-input");
  const bulkStepPreview = document.getElementById("bulk-step-preview");
  const bulkPreviewBody = document.getElementById("bulk-preview-body");
  const bulkPreviewCount = document.getElementById("bulk-preview-count");
  const bulkWarnings = document.getElementById("bulk-warnings");
  const btnBulkNext = document.getElementById("btn-bulk-next");
  const btnBulkConfirm = document.getElementById("btn-bulk-confirm");
  const btnBulkBack = document.getElementById("btn-bulk-back");
  const btnBulkCancel = document.getElementById("btn-bulk-cancel");

  let pendingBulkRecords = [];

  function resetBulkModal() {
    bulkTextarea.value = "";
    pendingBulkRecords = [];
    bulkPreviewBody.innerHTML = "";
    bulkWarnings.classList.add("hidden");
    bulkWarnings.innerHTML = "";
    // Show step 1, hide step 2
    bulkStepInput.classList.remove("hidden");
    bulkStepInput.classList.add("flex");
    bulkStepPreview.classList.add("hidden");
    bulkStepPreview.classList.remove("flex");
    // Show Next, hide Confirm & Back
    btnBulkNext.classList.remove("hidden");
    btnBulkConfirm.classList.add("hidden");
    btnBulkBack.classList.add("hidden");
  }

  // Open modal
  document.getElementById("btn-bulk-paste").addEventListener("click", () => {
    resetBulkModal();
    bulkModal.classList.remove("hidden");
    lucide.createIcons();
    // Focus the textarea after a brief delay for the animation
    setTimeout(() => bulkTextarea.focus(), 100);
  });

  // Cancel
  btnBulkCancel.addEventListener("click", () => {
    bulkModal.classList.add("hidden");
    resetBulkModal();
  });

  // Back to input step
  btnBulkBack.addEventListener("click", () => {
    bulkStepInput.classList.remove("hidden");
    bulkStepInput.classList.add("flex");
    bulkStepPreview.classList.add("hidden");
    bulkStepPreview.classList.remove("flex");
    btnBulkNext.classList.remove("hidden");
    btnBulkConfirm.classList.add("hidden");
    btnBulkBack.classList.add("hidden");
  });

  // Pre-visualizar
  btnBulkNext.addEventListener("click", () => {
    const raw = bulkTextarea.value.trim();

    if (!raw) {
      showToast("El área de texto está vacía. Pega tu lista primero.", "warning");
      return;
    }

    const lines = raw.split(/\r?\n/).filter(l => l.trim() !== "");

    if (lines.length === 0) {
      showToast("No se detectaron líneas válidas.", "warning");
      return;
    }

    // Read common fields from the main form
    const mainForm = document.getElementById("data-form");
    const fd = new FormData(mainForm);
    const commonEstado = (fd.get("estado") || "").trim().toUpperCase();
    const commonDestino = (fd.get("destino") || "").trim().toUpperCase();
    const commonRemision = (fd.get("remision") || "").trim().toUpperCase();
    const commonItem = (fd.get("item") || "").trim();
    const commonMaterial = (fd.get("material") || "").trim().toUpperCase();
    const commonFaltante = Number(fd.get("faltante")) || 0;
    const commonFechaSiniestro = fd.get("fecha_siniestro") || "";
    const commonConfronta1 = fd.get("confronta_1") || "";
    const commonFechaDe = (fd.get("fecha_de") || "").trim();
    const commonConfronta2 = (fd.get("confronta_2") || "").trim();

    // Calculate starting consecutivo
    let nextConsec = getNextConsecutivo();
    pendingBulkRecords = [];
    const skippedLines = [];

    lines.forEach((line, idx) => {
      // Split by tab or whitespace. The LAST token is the number, everything before is the marca.
      const parts = line.trim().split(/[\t\s]+/);

      if (parts.length < 2) {
        skippedLines.push({ lineNum: idx + 1, text: line.trim(), reason: "No se pudo separar marca y número" });
        return;
      }

      const numeroPart = parts[parts.length - 1];
      const marcaPart = parts.slice(0, -1).join(" ").toUpperCase();
      const numVal = Number(numeroPart);

      if (!marcaPart) {
        skippedLines.push({ lineNum: idx + 1, text: line.trim(), reason: "Marca vacía" });
        return;
      }

      if (isNaN(numVal)) {
        skippedLines.push({ lineNum: idx + 1, text: line.trim(), reason: `"${numeroPart}" no es un número válido` });
        return;
      }

      pendingBulkRecords.push({
        id_unico: Date.now().toString() + Math.random().toString(36).substr(2, 5),
        consecutivo: nextConsec++,
        estado: commonEstado,
        destino: commonDestino,
        marcas: marcaPart,
        numero: numVal,
        faltante: commonFaltante,
        fecha_siniestro: commonFechaSiniestro,
        confronta_1: commonConfronta1,
        fecha_de: commonFechaDe,
        confronta_2: commonConfronta2,
        remision: commonRemision,
        item: commonItem,
        material: commonMaterial,
        confrontada: false
      });
    });

    if (pendingBulkRecords.length === 0) {
      showToast("No se pudieron generar registros válidos. Verifica el formato.", "error");
      return;
    }

    // Render preview table
    bulkPreviewBody.innerHTML = "";
    pendingBulkRecords.forEach((r, i) => {
      const tr = document.createElement("tr");
      tr.className = "hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors";
      tr.innerHTML = `
        <td class="px-3 py-2 text-slate-400 dark:text-slate-500 font-medium">${i + 1}</td>
        <td class="px-3 py-2 font-semibold">${r.marcas}</td>
        <td class="px-3 py-2">${r.numero}</td>
        <td class="px-3 py-2 text-blue-600 dark:text-blue-400 font-semibold">${r.consecutivo}</td>
      `;
      bulkPreviewBody.appendChild(tr);
    });

    bulkPreviewCount.textContent = `Se van a agregar ${pendingBulkRecords.length} registro${pendingBulkRecords.length > 1 ? 's' : ''}`;

    // Show warnings for skipped lines
    if (skippedLines.length > 0) {
      bulkWarnings.classList.remove("hidden");
      bulkWarnings.innerHTML = `
        <div class="bg-amber-50 dark:bg-amber-900/20 text-amber-700 dark:text-amber-300 rounded-lg px-4 py-3 border border-amber-200 dark:border-amber-800 text-sm">
          <p class="font-semibold flex items-center gap-1.5 mb-1"><i data-lucide="alert-triangle" class="w-4 h-4"></i> ${skippedLines.length} línea${skippedLines.length > 1 ? 's' : ''} omitida${skippedLines.length > 1 ? 's' : ''}:</p>
          <ul class="list-disc list-inside text-xs space-y-0.5 mt-1">
            ${skippedLines.map(s => `<li>Línea ${s.lineNum}: <code class="font-mono bg-amber-100 dark:bg-amber-900/40 px-1 rounded">${s.text}</code> — ${s.reason}</li>`).join('')}
          </ul>
        </div>
      `;
    } else {
      bulkWarnings.classList.add("hidden");
      bulkWarnings.innerHTML = "";
    }

    // Switch to preview step
    bulkStepInput.classList.add("hidden");
    bulkStepInput.classList.remove("flex");
    bulkStepPreview.classList.remove("hidden");
    bulkStepPreview.classList.add("flex");
    btnBulkNext.classList.add("hidden");
    btnBulkConfirm.classList.remove("hidden");
    btnBulkBack.classList.remove("hidden");
    lucide.createIcons();
  });

  // Confirmar e Insertar
  btnBulkConfirm.addEventListener("click", () => {
    if (pendingBulkRecords.length === 0) return;

    STATE.records.push(...pendingBulkRecords);
    persistState();
    localStorage.removeItem("confrontas_manual_consecutivo");
    renderTable();
    renderSummary();
    consecutivoInput.value = getNextConsecutivo();

    const count = pendingBulkRecords.length;
    bulkModal.classList.add("hidden");
    resetBulkModal();

    showToast(`${count} registro${count > 1 ? 's' : ''} agregado${count > 1 ? 's' : ''} exitosamente vía Captura Rápida.`, "success");
  });
});

function getFilteredRecords() {
  if (!currentSearch) return STATE.records;
  
  return STATE.records.filter(r => {
    const fusionText = `${r.consecutivo} ${r.estado} ${r.destino} ${r.marcas} ${r.numero} ${r.remision} ${r.item} ${r.material}`.toLowerCase();
    return fusionText.includes(currentSearch);
  });
}

function renderTable() {
  const tbody = document.getElementById("table-body");
  tbody.innerHTML = "";
  
  const recordsToShow = getFilteredRecords();

  if (recordsToShow.length === 0) {
    if (STATE.records.length > 0 && currentSearch !== "") {
      tbody.innerHTML = `<tr><td colspan="14" class="px-6 py-8 text-center text-slate-500">Ningún registro coincide con la búsqueda "${currentSearch}".</td></tr>`;
    } else {
      tbody.innerHTML = `<tr><td colspan="14" class="px-6 py-8 text-center text-slate-400">No hay datos ingresados. Los registros aparecerán aquí.</td></tr>`;
    }
    return;
  }

  recordsToShow.forEach((r) => {
    const isConfrontada = !!r.confrontada;
    const tr = document.createElement("tr");
    tr.className = `transition-colors border-b border-slate-100 dark:border-slate-700/50 last:border-0 ${isConfrontada ? 'row-confrontada' : 'hover:bg-slate-50 dark:hover:bg-slate-700/50'}`;
    
    // Configuración de celdas: [field, label, type, isEditable, extraClasses]
    const cellConfig = [
        ["consecutivo", r.consecutivo, "number", false, "text-slate-500 dark:text-slate-400 font-medium"],
        ["estado", r.estado, "text", true, ""],
        ["destino", r.destino, "text", true, ""],
        ["marcas", r.marcas, "text", true, "font-bold opacity-90"],
        ["numero", r.numero, "number", true, "font-bold opacity-90"],
        ["faltante", Number(r.faltante).toFixed(2), "number", true, `text-red-600 dark:text-red-400 font-bold ${isConfrontada ? 'faltante-cell' : ''}`],
        ["fecha_siniestro", r.fecha_siniestro, "date", true, ""],
        ["confronta_1", r.confronta_1, "date", true, ""],
        ["fecha_de", r.fecha_de, "text", true, ""],
        ["confronta_2", r.confronta_2, "text", true, ""],
        ["remision", r.remision, "text", true, "font-semibold bg-blue-50/50 dark:bg-blue-900/20 text-blue-800 dark:text-blue-300 rounded"],
        ["item", r.item, "text", true, ""],
    ];

    cellConfig.forEach(([field, value, type, editable, classes]) => {
        const td = document.createElement("td");
        td.className = `px-4 py-3 whitespace-nowrap text-sm ${classes} ${editable ? 'editable-cell' : ''}`;
        
        if (editable) {
            td.innerHTML = `
                <div class="cell-display">
                    <span class="value-text">${value}</span>
                    <i data-lucide="edit-3" class="w-3 h-3 edit-hint"></i>
                </div>
            `;
            td.addEventListener("click", () => startEditing(td, r.id_unico, field, r[field], type));
        } else {
            td.textContent = value;
        }
        tr.appendChild(td);
    });

    // Columna M - Confronta / Acta de Hechos (Toggle Switch)
    const tdConfronta = document.createElement("td");
    tdConfronta.className = "px-4 py-3 whitespace-nowrap text-sm";
    tdConfronta.innerHTML = `
        <div class="flex items-center">
            <label class="confronta-switch" title="Marcar como confrontada">
                <input type="checkbox" ${isConfrontada ? 'checked' : ''} data-id="${r.id_unico}">
                <span class="confronta-slider"></span>
            </label>
            <span class="confronta-label ${isConfrontada ? 'text-emerald-600 dark:text-emerald-400' : 'text-slate-400 dark:text-slate-500'}">
                ${isConfrontada ? 'HECHO' : ''}
            </span>
        </div>
    `;
    const checkbox = tdConfronta.querySelector('input[type="checkbox"]');
    checkbox.addEventListener('change', (e) => {
        toggleConfronta(r.id_unico, e.target.checked);
    });
    tr.appendChild(tdConfronta);

    // Acción - Eliminar
    const tdAction = document.createElement("td");
    tdAction.className = "px-4 py-3 whitespace-nowrap text-sm text-right";
    tdAction.innerHTML = `
        <button onclick="removeRecord('${r.id_unico}')" class="text-rose-500 hover:text-rose-600 dark:text-rose-400 p-1" title="Eliminar">
            <i data-lucide="trash-2" class="w-4 h-4"></i>
        </button>
    `;
    tr.appendChild(tdAction);

    tbody.appendChild(tr);
  });
  
  lucide.createIcons();
}

/**
 * Inicia la edición de una celda
 */
function startEditing(td, idUnico, field, originalValue, type) {
    if (td.querySelector('input')) return; // Ya se está editando

    const displayDiv = td.querySelector('.cell-display');
    const valueText = displayDiv.querySelector('.value-text');
    
    const input = document.createElement("input");
    input.type = type === "date" ? "date" : (type === "number" ? "number" : "text");
    if (type === "number") input.step = "any";
    
    input.value = originalValue;
    input.className = "cell-input";
    
    // Ocultar contenido actual
    displayDiv.classList.add("hidden");
    td.appendChild(input);
    input.focus();
    if (type !== "date") input.select();

    const save = () => {
        let newValue = input.value;
        
        // Formateo según tipo
        if (type === "number") {
            newValue = Number(newValue);
            if (isNaN(newValue)) newValue = originalValue;
        } else if (type === "text") {
            // Mantener consistencia de mayúsculas si el valor original lo era (estética del app)
            if (originalValue === String(originalValue).toUpperCase()) {
                newValue = newValue.toUpperCase();
            }
            newValue = newValue.trim();
        }

        if (newValue !== originalValue) {
            const recordIdx = STATE.records.findIndex(rec => rec.id_unico === idUnico);
            if (recordIdx !== -1) {
                STATE.records[recordIdx][field] = newValue;
                persistState();
                
                // Si cambiamos algo que afecte el resumen o búsqueda, refrescamos todo
                // "Faltante" y "Remisión" afectan directamente al Resumen Vivo
                if (field === "faltante" || field === "remision") {
                    renderSummary();
                }
                
                // Actualizar el texto en la celda sin re-renderizar toda la fila para no perder el foco visual
                if (field === "faltante" && typeof newValue === 'number') {
                    valueText.textContent = newValue.toFixed(2);
                } else {
                    valueText.textContent = newValue;
                }
                
                // Si el campo es parte de la búsqueda activa, quizás convenga re-renderizar la tabla
                if (currentSearch && field !== "faltante") {
                    renderTable();
                    return;
                }
            }
        }

        // Restaurar vista
        input.remove();
        displayDiv.classList.remove("hidden");
    };

    input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") save();
        if (e.key === "Escape") {
            input.remove();
            displayDiv.classList.remove("hidden");
        }
    });

    input.addEventListener("blur", save);
}

/**
 * Toggle the confrontada state for a record
 */
window.toggleConfronta = function(idUnico, isChecked) {
  const idx = STATE.records.findIndex(r => r.id_unico === idUnico);
  if (idx !== -1) {
    STATE.records[idx].confrontada = isChecked;
    persistState();
    renderTable();
    renderSummary();
  }
};

function renderSummary() {
  const tbody = document.getElementById("summary-body");
  tbody.innerHTML = "";
  
  const recordsToSummarize = getFilteredRecords();
  // Separate totals: total vs pendiente (not confrontada)
  const summaryTotal = {};
  const summaryPendiente = {};

  recordsToSummarize.forEach(r => {
    const rem = r.remision || "SIN REMISIÓN";
    if (!summaryTotal[rem]) summaryTotal[rem] = 0;
    if (!summaryPendiente[rem]) summaryPendiente[rem] = 0;
    
    const val = Number(r.faltante);
    if (!isNaN(val)) {
      summaryTotal[rem] += val;
      // Solo sumar al pendiente si NO está confrontada
      if (!r.confrontada) {
        summaryPendiente[rem] += val;
      }
    }
  });

  const keys = Object.keys(summaryTotal);
  if (keys.length === 0) {
    tbody.innerHTML = `<tr><td colspan="2" class="px-4 py-3 text-center text-xs text-slate-400 dark:text-slate-500">Sin datos</td></tr>`;
    return;
  }

  keys.forEach(k => {
    const total = summaryTotal[k];
    const pendiente = summaryPendiente[k];
    const hasConfrontadas = total !== pendiente;
    
    const tr = document.createElement("tr");
    tr.className = "border-b border-slate-100 dark:border-slate-700/50 last:border-0 hover:bg-slate-50 dark:hover:bg-slate-700/50";
    tr.innerHTML = `
      <td class="px-4 py-3 text-sm font-medium opacity-90">${k}</td>
      <td class="px-4 py-3 text-right">
        <span class="text-sm text-red-600 dark:text-red-400 font-bold">${pendiente.toFixed(2)}</span>
        ${hasConfrontadas ? `<span class="block text-[0.65rem] text-slate-400 dark:text-slate-500 mt-0.5 line-through">${total.toFixed(2)} total</span>` : ''}
      </td>
    `;
    tbody.appendChild(tr);
  });
}

window.removeRecord = function(id) {
  try {
    const idx = STATE.records.findIndex(r => r.id_unico === id);
    if (idx !== -1) {
      STATE.records.splice(idx, 1);
      persistState();
      renderTable();
      renderSummary();
      showToast("Registro eliminado.", "warning");
    }
  } catch(e) {
    showToast("Error al eliminar el registro", "error");
  }
};

function showToast(msg, type = "success") {
  const t = document.createElement('div');
  
  const bgColors = {
    "success": "bg-emerald-500 dark:bg-emerald-600",
    "error": "bg-red-500 dark:bg-red-600",
    "warning": "bg-amber-500 dark:bg-amber-600"
  };
  
  const iconNames = {
    "success": "check-circle",
    "error": "alert-circle",
    "warning": "alert-triangle"
  };

  const bgColor = bgColors[type] || bgColors["success"];
  const iconName = iconNames[type] || iconNames["success"];

  t.className = `fixed bottom-4 right-4 z-[60] ${bgColor} text-white px-6 py-3 rounded-xl shadow-lg transform transition-all translate-y-0 opacity-100 font-medium flex items-center gap-2 border border-white/10`;
  t.innerHTML = `<i data-lucide="${iconName}" class="w-5 h-5"></i> <span>${msg}</span>`;
  
  document.body.appendChild(t);
  lucide.createIcons();
  
  setTimeout(() => {
    t.classList.add("translate-y-10", "opacity-0");
    setTimeout(() => t.remove(), 300);
  }, 4000); 
}

async function exportExcel() {
  try {
    if (STATE.records.length === 0) {
      showToast("No hay datos cargados para exportar.", "warning");
      return;
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = 'Confrontas 2026';
    wb.created = new Date();
    const ws = wb.addWorksheet('Concentrado', {
      views: [{ state: 'frozen', ySplit: 1 }]
    });

    // ── Column definitions with auto-width hints ──
    const colDefs = [
      { header: STATE.headers[0],  key: 'consecutivo',     width: 14 },
      { header: STATE.headers[1],  key: 'estado',          width: 16 },
      { header: STATE.headers[2],  key: 'destino',         width: 12 },
      { header: STATE.headers[3],  key: 'marcas',          width: 12 },
      { header: STATE.headers[4],  key: 'numero',          width: 12 },
      { header: STATE.headers[5],  key: 'faltante',        width: 16 },
      { header: STATE.headers[6],  key: 'fecha_siniestro', width: 16 },
      { header: STATE.headers[7],  key: 'confronta_1',     width: 16 },
      { header: STATE.headers[8],  key: 'fecha_de',        width: 14 },
      { header: STATE.headers[9],  key: 'confronta_2',     width: 16 },
      { header: STATE.headers[10], key: 'remision',        width: 16 },
      { header: STATE.headers[11], key: 'item',            width: 12 },
      { header: STATE.headers[12], key: 'confronta_acta',  width: 26 },
      { header: '',                key: 'spacer',          width: 4  },
      { header: 'Remisión',        key: 'sum_remision',    width: 20 },
      { header: 'Total Faltante',  key: 'sum_faltante',    width: 18 },
    ];
    ws.columns = colDefs;

    // ── Styles ──
    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };
    const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10, name: 'Calibri' };
    const headerBorder = {
      top:    { style: 'thin', color: { argb: 'FF334155' } },
      bottom: { style: 'medium', color: { argb: 'FF1E3A5F' } },
      left:   { style: 'thin', color: { argb: 'FF334155' } },
      right:  { style: 'thin', color: { argb: 'FF334155' } },
    };
    const cellBorder = {
      top:    { style: 'thin', color: { argb: 'FFE2E8F0' } },
      bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      left:   { style: 'thin', color: { argb: 'FFE2E8F0' } },
      right:  { style: 'thin', color: { argb: 'FFE2E8F0' } },
    };
    const faltanteHeaderFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDC2626' } };
    const hechoFill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
    const hechoFont   = { bold: true, color: { argb: 'FF166534' }, size: 10, name: 'Calibri' };
    const pendienteFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF7ED' } };
    const pendienteFont = { color: { argb: 'FF9A3412' }, size: 10, name: 'Calibri' };
    const faltanteCellFont = { bold: true, color: { argb: 'FFDC2626' }, size: 10, name: 'Calibri' };
    const summaryHeaderFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF059669' } };
    const summaryHeaderFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10, name: 'Calibri' };
    const dataFont = { size: 10, name: 'Calibri', color: { argb: 'FF1E293B' } };
    const dateFormat = 'DD/MM/YYYY';

    // ── Style header row ──
    const headerRow = ws.getRow(1);
    headerRow.height = 28;
    headerRow.eachCell((cell, colNumber) => {
      if (colNumber <= 13) {
        cell.fill = colNumber === 6 ? faltanteHeaderFill : headerFill;
        cell.font = headerFont;
        cell.border = headerBorder;
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      } else if (colNumber >= 15) {
        cell.fill = summaryHeaderFill;
        cell.font = summaryHeaderFont;
        cell.border = headerBorder;
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    });

    // ── Add data rows ──
    STATE.records.forEach((r, idx) => {
      const fechaSiniestro = r.fecha_siniestro ? new Date(r.fecha_siniestro + 'T00:00:00') : '';
      const confronta1 = r.confronta_1 ? new Date(r.confronta_1 + 'T00:00:00') : '';
      const isHecho = !!r.confrontada;

      const dataRow = ws.addRow({
        consecutivo:    Number(r.consecutivo) || null,
        estado:         r.estado,
        destino:        r.destino,
        marcas:         r.marcas,
        numero:         Number(r.numero) || null,
        faltante:       Number(r.faltante) || 0,
        fecha_siniestro: fechaSiniestro,
        confronta_1:    confronta1,
        fecha_de:       r.fecha_de,
        confronta_2:    r.confronta_2,
        remision:       r.remision,
        item:           r.item,
        confronta_acta: isHecho ? 'HECHO' : '',
      });

      dataRow.height = 22;

      // Style each data cell
      dataRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber > 13) return;

        cell.font = dataFont;
        cell.border = cellBorder;
        cell.alignment = { vertical: 'middle' };

        // Alternate row shading
        if (idx % 2 === 1) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        }

        // Faltante column (6) — red bold
        if (colNumber === 6) {
          cell.font = faltanteCellFont;
          cell.numFmt = '#,##0.00';
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
          if (isHecho) {
            // Struck-through faltante for confronted rows
            cell.font = { ...faltanteCellFont, strike: true, color: { argb: 'FF94A3B8' } };
          }
        }

        // Date columns (7, 8)
        if (colNumber === 7 || colNumber === 8) {
          if (cell.value instanceof Date) {
            cell.numFmt = dateFormat;
          }
        }

        // Confronta / Acta column (13)
        if (colNumber === 13) {
          if (isHecho) {
            cell.fill = hechoFill;
            cell.font = hechoFont;
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
          } else {
            cell.fill = pendienteFill;
            cell.font = pendienteFont;
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
          }
        }

        // Remisión column (11) — subtle blue
        if (colNumber === 11) {
          cell.font = { ...dataFont, bold: true, color: { argb: 'FF1E40AF' } };
        }

        // Consecutivo column (1)
        if (colNumber === 1) {
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.font = { ...dataFont, bold: true, color: { argb: 'FF3B82F6' } };
        }

        // Marcas & Numero (4, 5) — bold
        if (colNumber === 4 || colNumber === 5) {
          cell.font = { ...dataFont, bold: true };
        }
      });

      // If entire row is confrontada, give it full green tint
      if (isHecho) {
        for (let c = 1; c <= 13; c++) {
          const cell = dataRow.getCell(c);
          if (c !== 6 && c !== 13) { // Keep faltante and acta with their own styles
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0FDF4' } };
          }
        }
      }
    });

    // ── SUMIF Summary (columns O-P) ──
    const summary = {};
    STATE.records.forEach(r => {
      const rem = r.remision || 'SIN REMISIÓN';
      if (!summary[rem]) summary[rem] = 0;
      if (!r.confrontada) {
        summary[rem] += (Number(r.faltante) || 0);
      }
    });

    const uniqueRemisiones = Object.keys(summary);
    uniqueRemisiones.forEach((rem, idx) => {
      const rowNum = idx + 2; // Row 2 onwards (after header)
      const row = ws.getRow(rowNum);
      
      const cellO = row.getCell(15);
      cellO.value = rem;
      cellO.font = { bold: true, size: 10, name: 'Calibri', color: { argb: 'FF1E293B' } };
      cellO.border = cellBorder;
      cellO.alignment = { vertical: 'middle' };

      const cellP = row.getCell(16);
      cellP.value = summary[rem];
      cellP.numFmt = '#,##0.00';
      cellP.font = { bold: true, size: 11, name: 'Calibri', color: { argb: 'FFDC2626' } };
      cellP.border = cellBorder;
      cellP.alignment = { vertical: 'middle', horizontal: 'right' };
      cellP.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF1F2' } };
    });

    // ── Auto-fit column widths based on content ──
    ws.columns.forEach((col) => {
      let maxLen = col.header ? col.header.length : 0;
      col.eachCell({ includeEmpty: false }, (cell) => {
        const cellLen = cell.value ? String(cell.value).length : 0;
        if (cellLen > maxLen) maxLen = cellLen;
      });
      // Clamp between original width and calculated, with padding
      col.width = Math.max(col.width || 10, Math.min(maxLen + 4, 40));
    });

    // ── Generate and download ──
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Concentrado_Confrontas_Web_Export.xlsx';
    a.click();
    URL.revokeObjectURL(url);

    showToast("Archivo Excel con estilos generado exitosamente.", "success");

  } catch (error) {
    console.error("Error crítico durante la exportación a Excel", error);
    showToast("Error crítico en la exportación: " + error.message, "error");
  }
}
