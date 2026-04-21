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
    "Nombre Material SAP"
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
  
  function applyTheme(isDark) {
    if (isDark) {
      document.documentElement.classList.add('dark');
      themeIconLight.classList.remove('hidden');
      themeIconDark.classList.add('hidden');
    } else {
      document.documentElement.classList.remove('dark');
      themeIconLight.classList.add('hidden');
      themeIconDark.classList.remove('hidden');
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

  document.getElementById('theme-toggle').addEventListener('click', () => {
    const currentlyDark = document.documentElement.classList.contains('dark');
    const newThemeDark = !currentlyDark;
    localStorage.setItem('theme', newThemeDark ? 'dark' : 'light');
    applyTheme(newThemeDark);
  });
}

document.addEventListener("DOMContentLoaded", () => {
  lucide.createIcons();
  const form = document.getElementById("data-form");
  const consecutivoInput = document.getElementById("consecutivo");

  initTheme();

  // Arrancar estado
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
        material: fd.get("material").trim().toUpperCase()
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
             material: String(row[12] || "").trim().toUpperCase()
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
      tbody.innerHTML = `<tr><td colspan="13" class="px-6 py-8 text-center text-slate-500">Ningún registro coincide con la búsqueda "${currentSearch}".</td></tr>`;
    } else {
      tbody.innerHTML = `<tr><td colspan="13" class="px-6 py-8 text-center text-slate-400">No hay datos ingresados. Los registros aparecerán aquí.</td></tr>`;
    }
    return;
  }

  recordsToShow.forEach((r) => {
    const tr = document.createElement("tr");
    tr.className = "hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors border-b border-slate-100 dark:border-slate-700/50 last:border-0";
    tr.innerHTML = `
      <td class="px-4 py-3 whitespace-nowrap text-sm text-slate-500 dark:text-slate-400 font-medium">${r.consecutivo}</td>
      <td class="px-4 py-3 whitespace-nowrap text-sm">${r.estado}</td>
      <td class="px-4 py-3 whitespace-nowrap text-sm">${r.destino}</td>
      <td class="px-4 py-3 whitespace-nowrap text-sm font-bold opacity-90">${r.marcas} ${r.numero}</td>
      <td class="px-4 py-3 whitespace-nowrap text-sm text-red-600 dark:text-red-400 font-bold">${Number(r.faltante).toFixed(2)}</td>
      <td class="px-4 py-3 whitespace-nowrap text-sm">${r.fecha_siniestro}</td>
      <td class="px-4 py-3 whitespace-nowrap text-sm">${r.confronta_1}</td>
      <td class="px-4 py-3 whitespace-nowrap text-sm">${r.fecha_de}</td>
      <td class="px-4 py-3 whitespace-nowrap text-sm">${r.confronta_2}</td>
      <td class="px-4 py-3 whitespace-nowrap text-sm font-semibold bg-blue-50/50 dark:bg-blue-900/20 text-blue-800 dark:text-blue-300 rounded">${r.remision}</td>
      <td class="px-4 py-3 whitespace-nowrap text-sm">${r.item}</td>
      <td class="px-4 py-3 whitespace-nowrap text-sm">${r.material}</td>
      <td class="px-4 py-3 whitespace-nowrap text-sm text-right">
        <!-- Eliminamos por ID interno para no cruzar indices en vista filtrada -->
        <button onclick="removeRecord('${r.id_unico}')" class="text-rose-500 hover:text-rose-600 dark:text-rose-400 p-1" title="Eliminar"><i data-lucide="trash-2" class="w-4 h-4"></i></button>
      </td>
    `;
    tbody.appendChild(tr);
  });
  
  lucide.createIcons();
}

function renderSummary() {
  const tbody = document.getElementById("summary-body");
  tbody.innerHTML = "";
  
  const recordsToSummarize = getFilteredRecords();
  const summary = {};

  recordsToSummarize.forEach(r => {
    const rem = r.remision || "SIN REMISIÓN";
    if (!summary[rem]) summary[rem] = 0;
    
    const val = Number(r.faltante);
    if (!isNaN(val)) {
      summary[rem] += val;
    }
  });

  const keys = Object.keys(summary);
  if (keys.length === 0) {
    tbody.innerHTML = `<tr><td colspan="2" class="px-4 py-3 text-center text-xs text-slate-400 dark:text-slate-500">Sin datos</td></tr>`;
    return;
  }

  keys.forEach(k => {
    const tr = document.createElement("tr");
    tr.className = "border-b border-slate-100 dark:border-slate-700/50 last:border-0 hover:bg-slate-50 dark:hover:bg-slate-700/50";
    tr.innerHTML = `
      <td class="px-4 py-3 text-sm font-medium opacity-90">${k}</td>
      <td class="px-4 py-3 text-sm text-red-600 dark:text-red-400 font-bold text-right">${summary[k].toFixed(2)}</td>
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

function exportExcel() {
  try {
    // Es mejor exportar SIEMPRE todos los records de la sesión, no solo los filtrados visualmente.
    if (STATE.records.length === 0) {
      showToast("No hay datos cargados para exportar.", "warning");
      return;
    }

    const aoa = [STATE.headers];

    STATE.records.forEach(r => {
      const row = [
        Number(r.consecutivo) || null, 
        r.estado,                      
        r.destino,                     
        r.marcas,                      
        Number(r.numero) || null,      
        Number(r.faltante) || 0,       
        dateToExcelCell(r.fecha_siniestro), 
        dateToExcelCell(r.confronta_1),     
        r.fecha_de,                    
        r.confronta_2,                 
        r.remision,                    
        r.item,                        
        r.material                     
      ];
      aoa.push(row);
    });

    aoa[0][14] = "Remisión";      
    aoa[0][15] = "Total Faltante";
    
    // El resumen en excel se basa en la hoja entera sin filtro
    const summary = {};
    STATE.records.forEach(r => {
      const rem = r.remision || "SIN REMISIÓN";
      if (!summary[rem]) summary[rem] = 0;
      summary[rem] += (Number(r.faltante) || 0);
    });

    const uniqueRemisiones = Object.keys(summary);
    uniqueRemisiones.forEach((rem, idx) => {
      const rIndex = idx + 1;
      
      if (!aoa[rIndex]) {
        aoa[rIndex] = [];
      }
      
      for (let i = aoa[rIndex].length; i < 14; i++) {
        aoa[rIndex][i] = null;
      }
      
      aoa[rIndex][14] = rem;           
      aoa[rIndex][15] = summary[rem];  
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    
    ws['!cols'] = [
      {wch: 12}, // A
      {wch: 15}, // B
      {wch: 10}, // C
      {wch: 10}, // D
      {wch: 10}, // E
      {wch: 15}, // F
      {wch: 15}, // G
      {wch: 15}, // H
      {wch: 15}, // I
      {wch: 15}, // J
      {wch: 15}, // K
      {wch: 10}, // L
      {wch: 25}, // M
      {wch: 5},  // N 
      {wch: 20}, // O
      {wch: 18}, // P
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, "Concentrado_Confrontas_Web_Export.xlsx");
    
    showToast("Archivo Excel .xlsx generado exitosamente.", "success");

  } catch (error) {
    console.error("Error crítico durante la exportación a Excel", error);
    showToast("Error crítico en la exportación.", "error");
  }
}
