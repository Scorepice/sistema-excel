/* Dashboard interactive department selector + progressive reveal */
(function(){
  function updateDashboardCounters(totalVisible, selectedCount) {
    const visibleEl = document.getElementById('dashboardDeptCount');
    const selectedEl = document.getElementById('dashboardSelectedLabel');
    if (visibleEl) visibleEl.textContent = `${totalVisible} visibles`;
    if (selectedEl) selectedEl.textContent = `${selectedCount} departamentos`;
  }

  function updateEstadoCounters(totalVisible, selectedCount) {
    const visibleEl = document.getElementById('dashboardStateCount');
    const selectedEl = document.getElementById('dashboardStateSelectedLabel');
    if (visibleEl) visibleEl.textContent = `${totalVisible} visibles`;
    if (selectedEl) selectedEl.textContent = `${selectedCount} estados`;
  }

  function updateVistaLabel(label) {
    const vistaEl = document.getElementById('dashboardVistaLabel');
    if (vistaEl) vistaEl.textContent = label;
  }

  // Previously clicking departments forced 'depto' view; avoid that now.
  // Keep this helper for backward compatibility but it will not be used by clicks.
  function forceDeptoView() {
    const filtro = document.getElementById('filtroDatos');
    if (filtro && filtro.value !== 'depto') {
      filtro.value = 'depto';
    }
    if (typeof actualizarGrafica === 'function') {
      actualizarGrafica();
    }
  }

  function createDeptSelector(dataDepto, rootId='deptSelectorRoot'){
    const root = document.getElementById(rootId);
    if(!root) return;
    root.innerHTML = '';

    const container = document.createElement('div');
    container.className = 'dept-selector-root';

    const input = document.createElement('input');
    input.type = 'search';
    input.placeholder = 'Buscar departamento...';
    input.className = 'dept-search';
    container.appendChild(input);

    const list = document.createElement('div');
    list.className = 'dept-list';
    container.appendChild(list);

    const controls = document.createElement('div');
    controls.className = 'dept-actions';
    const btnAll = document.createElement('button'); btnAll.type = 'button'; btnAll.className='btn btn-sm btn-primary'; btnAll.textContent='Seleccionar todos';
    const btnNone = document.createElement('button'); btnNone.type = 'button'; btnNone.className='btn btn-sm btn-outline-secondary'; btnNone.textContent='Limpiar';
    controls.appendChild(btnAll); controls.appendChild(btnNone);
    container.appendChild(controls);

    const maxValue = Math.max(...dataDepto.valores, 1);
    dataDepto.labels.forEach(function(label, idx){
      const card = document.createElement('button');
      card.type = 'button';
      card.className = 'dept-chip is-dimmed';
      card.dataset.idx = idx;
      const count = dataDepto.valores[idx] || 0;
      const fillWidth = Math.max(12, Math.round((count / maxValue) * 100));
      const color = dataDepto.colores[idx] || '#17659d';
      card.innerHTML = `
        <div class="dept-chip-header">
          <div>
            <div class="dept-chip-name">${label}</div>
            <div class="dept-chip-subtext">Haz click para activarlo en la gráfica</div>
          </div>
          <div class="dept-chip-count">${count}</div>
        </div>
        <div class="dept-chip-track"><div class="dept-chip-fill" style="width:${fillWidth}%; background: linear-gradient(90deg, ${color}, rgba(255,255,255,0.8));"></div></div>
      `;
      card.addEventListener('click', function(){
        card.classList.toggle('is-selected');
        card.classList.toggle('is-dimmed', !card.classList.contains('is-selected'));
        // Preserve current filtro (view) and update the chart to reflect selection.
        if(typeof actualizarGrafica === 'function') actualizarGrafica();
      });
      list.appendChild(card);
    });

    input.addEventListener('input', function(){
      const q = input.value.toLowerCase();
      Array.from(list.children).forEach(function(p){
        const text = p.textContent.toLowerCase();
        p.style.display = text.indexOf(q) !== -1 ? 'block' : 'none';
      });
      updateDashboardCounters(Array.from(list.children).filter(p => p.style.display !== 'none').length, getSelectedDeptIndices(rootId).length);
    });

    btnAll.addEventListener('click', function(){
      Array.from(list.children).forEach(p=>{
        p.classList.add('is-selected');
        p.classList.remove('is-dimmed');
      });
      if(typeof actualizarGrafica === 'function') actualizarGrafica();
    });
    btnNone.addEventListener('click', function(){
      Array.from(list.children).forEach(p=>{
        p.classList.remove('is-selected');
        p.classList.add('is-dimmed');
      });
      if(typeof actualizarGrafica === 'function') actualizarGrafica();
    });

    root.appendChild(container);
    updateDashboardCounters(dataDepto.labels.length, getSelectedDeptIndices(rootId).length);
  }

  function createStateSelector(dataEstado, rootId='stateSelectorRoot'){
    const root = document.getElementById(rootId);
    if(!root) return;
    root.innerHTML = '';

    const container = document.createElement('div');
    container.className = 'state-selector-root';

    const list = document.createElement('div');
    list.className = 'state-list';
    container.appendChild(list);

    const controls = document.createElement('div');
    controls.className = 'state-actions dept-actions';
    const btnAll = document.createElement('button');
    btnAll.type = 'button';
    btnAll.className = 'btn btn-sm btn-primary';
    btnAll.textContent = 'Seleccionar todos';
    const btnNone = document.createElement('button');
    btnNone.type = 'button';
    btnNone.className = 'btn btn-sm btn-outline-secondary';
    btnNone.textContent = 'Limpiar';
    controls.appendChild(btnAll);
    controls.appendChild(btnNone);
    container.appendChild(controls);

    const labels = Array.isArray(dataEstado.total_labels) ? dataEstado.total_labels : [];
    const valores = Array.isArray(dataEstado.total_valores) ? dataEstado.total_valores : [];
    const colores = Array.isArray(dataEstado.total_colores) ? dataEstado.total_colores : [];
    const maxValue = Math.max(...valores, 1);

    labels.forEach(function(label, idx){
      const card = document.createElement('button');
      card.type = 'button';
      card.className = 'dept-chip estado-chip is-dimmed';
      card.dataset.idx = idx;
      const count = valores[idx] || 0;
      const fillWidth = Math.max(12, Math.round((count / maxValue) * 100));
      const color = colores[idx] || '#17659d';
      card.innerHTML = `
        <div class="dept-chip-header">
          <div>
            <div class="dept-chip-name">${label}</div>
            <div class="dept-chip-subtext">Haz click para filtrar por este estado</div>
          </div>
          <div class="dept-chip-count">${count}</div>
        </div>
        <div class="dept-chip-track"><div class="dept-chip-fill" style="width:${fillWidth}%; background: linear-gradient(90deg, ${color}, rgba(255,255,255,0.8));"></div></div>
      `;
      card.addEventListener('click', function(){
        card.classList.toggle('is-selected');
        card.classList.toggle('is-dimmed', !card.classList.contains('is-selected'));
        if(typeof actualizarGrafica === 'function') actualizarGrafica();
      });
      list.appendChild(card);
    });

    btnAll.addEventListener('click', function(){
      Array.from(list.children).forEach(p=>{
        p.classList.add('is-selected');
        p.classList.remove('is-dimmed');
      });
      if(typeof actualizarGrafica === 'function') actualizarGrafica();
    });

    btnNone.addEventListener('click', function(){
      Array.from(list.children).forEach(p=>{
        p.classList.remove('is-selected');
        p.classList.add('is-dimmed');
      });
      if(typeof actualizarGrafica === 'function') actualizarGrafica();
    });

    root.appendChild(container);
    updateEstadoCounters(labels.length, getSelectedStateIndices(rootId).length);
  }

  function getSelectedDeptIndices(rootId='deptSelectorRoot'){
    const root = document.getElementById(rootId);
    if(!root) return [];
    const pills = root.querySelectorAll('.dept-chip');
    const sel = [];
    pills.forEach(function(p){ if(p.classList.contains('is-selected')) sel.push(Number(p.dataset.idx)); });
    return sel;
  }

  function getSelectedStateIndices(rootId='stateSelectorRoot'){
    const root = document.getElementById(rootId);
    if(!root) return [];
    const pills = root.querySelectorAll('.estado-chip');
    const sel = [];
    pills.forEach(function(p){ if(p.classList.contains('is-selected')) sel.push(Number(p.dataset.idx)); });
    return sel;
  }

  async function animateDeptSelection(chartInstance, labels, values, colors){
    // progressive reveal: clear then add one by one
    chartInstance.data.labels = [];
    if(chartInstance.data.datasets && chartInstance.data.datasets.length>0){
      chartInstance.data.datasets[0].data = [];
      chartInstance.data.datasets[0].backgroundColor = [];
    } else {
      chartInstance.data.datasets = [{ label: 'Total RDM', data: [], backgroundColor: [] }];
    }
    chartInstance.update();

    for(let i=0;i<labels.length;i++){
      chartInstance.data.labels.push(labels[i]);
      chartInstance.data.datasets[0].data.push(values[i]);
      chartInstance.data.datasets[0].backgroundColor.push(colors[i] || '#6c757d');
      chartInstance.update();
      // small delay for visual effect
      await new Promise(r=>setTimeout(r, 180));
    }
  }

  async function animateAddDatasets(chartInstance, datasets){
    // assumes chartInstance.data.labels already set
    chartInstance.data.datasets = [];
    chartInstance.update();
    for(let i=0;i<datasets.length;i++){
      chartInstance.data.datasets.push(datasets[i]);
      chartInstance.update();
      await new Promise(r=>setTimeout(r, 220));
    }
  }

  // Expose helpers
  window.createDeptSelector = createDeptSelector;
  window.createStateSelector = createStateSelector;
  window.getSelectedDeptIndices = getSelectedDeptIndices;
  window.getSelectedStateIndices = getSelectedStateIndices;
  window.animateDeptSelection = animateDeptSelection;
  window.updateVistaLabel = updateVistaLabel;
  window.forceDeptoView = forceDeptoView;
  window.animateAddDatasets = animateAddDatasets;

})();
