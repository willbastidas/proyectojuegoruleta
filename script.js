// --- VARIABLES GLOBALES ---
let listaParticipantes = [];
let listaJuegos = [];
let dueloActual = null;
let rotacionAcumulada = 0;
let juegosDisponibles = [];
let contadorPartidas = 1;

const coloresRuleta = ['#f1c40f', '#e74c3c', '#3498db', '#2ecc71', '#9b59b6', '#e67e22', '#1abc9c', '#34495e'];

// --- ELEMENTOS DEL DOM ---
const cantidadParticipantes = document.getElementById('cantidad-participantes');
const radiosIngreso = document.getElementsByName('metodo-ingreso');
const contenedorManual = document.getElementById('contenedor-manual');
const contenedorExcel = document.getElementById('contenedor-excel');
const listaInputsManual = document.getElementById('lista-inputs-manual');
const archivoExcel = document.getElementById('archivo-excel');

const cantidadJuegos = document.getElementById('cantidad-juegos');
const radiosJuego = document.getElementsByName('metodo-juego');
const contenedorJuegoManual = document.getElementById('contenedor-juego-manual');
const contenedorJuegoExcel = document.getElementById('contenedor-juego-excel');
const listaInputsJuegos = document.getElementById('lista-inputs-juegos');
const archivoExcelJuegos = document.getElementById('archivo-excel-juegos');

const tipoFormato = document.getElementById('tipo-formato');
const btnIniciar = document.getElementById('btn-iniciar');
const seccionConfig = document.getElementById('configuracion');

const seccionGrupos = document.getElementById('seccion-grupos');
const contenedorGrupos = document.getElementById('contenedor-grupos');
const btnIrBracket = document.getElementById('btn-ir-bracket');

const seccionTorneo = document.getElementById('seccion-torneo');
const contenedorTorneo = document.getElementById('contenedor-torneo');
const listaHistorico = document.getElementById('lista-historico');

const modalRuleta = document.getElementById('modal-ruleta');
const btnCerrarModal = document.getElementById('cerrar-modal');
const canvasRuleta = document.getElementById('canvas-ruleta');
const ctx = canvasRuleta ? canvasRuleta.getContext('2d') : null;
const btnGirar = document.getElementById('btn-girar');
const ruletaPantalla = document.getElementById('ruleta-pantalla');
const juegoAsignado = document.getElementById('juego-asignado');

// --- 1. LÓGICA DE INTERFAZ ---
function generarInputsManuales() {
    const cantidad = parseInt(cantidadParticipantes.value) || 0;
    listaInputsManual.innerHTML = '';
    for (let i = 1; i <= cantidad; i++) {
        const input = document.createElement('input');
        input.type = 'text';
        input.placeholder = `Jugador ${i}`;
        input.className = 'input-jugador';
        listaInputsManual.appendChild(input);
    }
}

function generarInputsJuegos() {
    const cantidad = parseInt(cantidadJuegos.value) || 0;
    listaInputsJuegos.innerHTML = '';
    for (let i = 1; i <= cantidad; i++) {
        const input = document.createElement('input');
        input.type = 'text';
        input.placeholder = `Juego ${i}`;
        input.className = 'input-juego';
        listaInputsJuegos.appendChild(input);
    }
}

cantidadParticipantes.addEventListener('input', generarInputsManuales);
cantidadJuegos.addEventListener('input', generarInputsJuegos);
generarInputsManuales();
generarInputsJuegos();

radiosIngreso.forEach(radio => radio.addEventListener('change', (e) => {
    contenedorManual.classList.toggle('oculto', e.target.value !== 'manual');
    contenedorExcel.classList.toggle('oculto', e.target.value === 'manual');
}));

radiosJuego.forEach(radio => radio.addEventListener('change', (e) => {
    contenedorJuegoManual.classList.toggle('oculto', e.target.value !== 'manual');
    contenedorJuegoExcel.classList.toggle('oculto', e.target.value === 'manual');
}));

function leerArchivoExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const filas = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            resolve(filas.map(f => f[0]).filter(v => v && String(v).trim() !== '').map(String));
        };
        reader.onerror = error => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

// --- 2. INICIAR EL TORNEO ---
btnIniciar.addEventListener('click', async () => {
    listaParticipantes = [];
    listaJuegos = [];

    if (document.querySelector('input[name="metodo-ingreso"]:checked').value === 'manual') {
        document.querySelectorAll('.input-jugador').forEach(i => { if (i.value.trim()) listaParticipantes.push(i.value.trim()); });
    } else {
        const file = archivoExcel.files[0];
        if (!file) return alert('Por favor, selecciona el archivo Excel de Jugadores.');
        listaParticipantes = await leerArchivoExcel(file);
    }

    if (document.querySelector('input[name="metodo-juego"]:checked').value === 'manual') {
        document.querySelectorAll('.input-juego').forEach(i => { if (i.value.trim()) listaJuegos.push(i.value.trim()); });
    } else {
        const fileJuegos = archivoExcelJuegos.files[0];
        if (!fileJuegos) return alert('Por favor, selecciona el archivo Excel de Juegos.');
        listaJuegos = await leerArchivoExcel(fileJuegos);
    }

    if (listaParticipantes.length < 3) return alert(`Solo detectamos ${listaParticipantes.length} participantes. Ingresa al menos 3.`);
    if (listaJuegos.length === 0) return alert('Por favor, ingresa al menos 1 juego.');

    const formatoElegido = tipoFormato.value;
    if (formatoElegido === 'grupos' && listaParticipantes.length < 8) {
        return alert("Para crear una Fase de Grupos necesitas al menos 8 participantes.");
    }

    juegosDisponibles = [...listaJuegos];
    contadorPartidas = 1;
    listaHistorico.innerHTML = '';
    seccionConfig.classList.add('oculto');

    if (formatoElegido === 'grupos') {
        seccionGrupos.classList.remove('oculto');
        seccionTorneo.classList.remove('oculto'); // Mostramos el historial abajo
        document.getElementById('titulo-formato').textContent = "Esperando Seleccionados para el Bracket...";
        generarFaseGrupos();
    } else {
        seccionTorneo.classList.remove('oculto');
        generarBracket(listaParticipantes);
    }
});

// --- 3. LÓGICA DE FASE DE GRUPOS ---
function generarFaseGrupos() {
    contenedorGrupos.innerHTML = '';

    // Mezclar aleatoriamente
    let mezclados = [...listaParticipantes].sort(() => Math.random() - 0.5);

    // Calcular cantidad de grupos (aprox 4 personas por grupo)
    let numGrupos = Math.max(2, Math.floor(mezclados.length / 4));
    let grupos = Array.from({ length: numGrupos }, () => []);

    // Repartir jugadores en los grupos
    mezclados.forEach((jugador, i) => grupos[i % numGrupos].push(jugador));

    grupos.forEach((grupo, gIndex) => {
        const card = document.createElement('div');
        card.className = 'grupo-card';
        card.innerHTML = `<h3 class="grupo-title">Grupo ${String.fromCharCode(65 + gIndex)}</h3>`;

        const divPartidos = document.createElement('div');
        divPartidos.className = 'grupo-partidos';

        // Generar partidos Todos contra Todos (Round Robin)
        for (let i = 0; i < grupo.length; i++) {
            for (let j = i + 1; j < grupo.length; j++) {
                const matchBox = crearCajaMatch(grupo[i], grupo[j], false);
                divPartidos.appendChild(matchBox);
            }
        }
        card.appendChild(divPartidos);

        // Generar zona de selección manual de clasificados
        const divSeleccion = document.createElement('div');
        divSeleccion.className = 'grupo-seleccion';
        divSeleccion.innerHTML = `<strong style="color:#2c3e50;">✔ Selecciona los que avanzan:</strong>`;

        grupo.forEach(jugador => {
            const label = document.createElement('label');
            label.innerHTML = `<input type="checkbox" class="chk-avanza" value="${jugador}"> ${jugador}`;
            divSeleccion.appendChild(label);
        });

        card.appendChild(divSeleccion);
        contenedorGrupos.appendChild(card);
    });
}

// Botón para finalizar grupos y pasar al bracket
btnIrBracket.addEventListener('click', () => {
    const checkboxes = document.querySelectorAll('.chk-avanza:checked');
    const clasificados = Array.from(checkboxes).map(chk => chk.value);

    if (clasificados.length < 2) {
        return alert("Debes seleccionar al menos a 2 jugadores para armar el bracket final.");
    }

    // Ocultar sección de grupos y dibujar el bracket con los elegidos
    seccionGrupos.classList.add('oculto');
    document.getElementById('titulo-formato').textContent = `Bracket Final (${clasificados.length} Clasificados)`;
    generarBracket(clasificados);
});

// --- 4. FUNCIÓN REUTILIZABLE PARA CREAR CAJAS DE MATCH ---
function crearCajaMatch(jugador1Nombre, jugador2Nombre, esBracket, funcionAvanzarJ1 = null, funcionAvanzarJ2 = null) {
    const divEnfrentamiento = document.createElement('div');
    divEnfrentamiento.className = 'enfrentamiento';
    divEnfrentamiento.style.margin = esBracket ? "15px 0" : "15px 5px 30px 5px"; // Más espacio si es de grupos

    // --- JUGADOR 1 ---
    const j1 = document.createElement('div');
    j1.className = 'jugador';
    const spanJ1 = document.createElement('span');
    spanJ1.className = 'nombre';
    spanJ1.textContent = jugador1Nombre;

    const divBotones1 = document.createElement('div');
    divBotones1.className = 'botones-container';

    const btnJuego1 = document.createElement('button');
    btnJuego1.className = 'btn-win-juego';
    btnJuego1.textContent = '🎮 Win Juego';

    const btnPartida1 = document.createElement('button');
    btnPartida1.className = 'btn-win-partida';
    btnPartida1.textContent = esBracket ? '🏆 Win Partida' : '✅ Fin Match';

    divBotones1.appendChild(btnJuego1);
    divBotones1.appendChild(btnPartida1);
    j1.appendChild(spanJ1);
    j1.appendChild(divBotones1);

    // --- JUGADOR 2 ---
    const j2 = document.createElement('div');
    j2.className = 'jugador';
    const spanJ2 = document.createElement('span');
    spanJ2.className = 'nombre';
    spanJ2.textContent = jugador2Nombre;

    const divBotones2 = document.createElement('div');
    divBotones2.className = 'botones-container';

    const btnJuego2 = document.createElement('button');
    btnJuego2.className = 'btn-win-juego';
    btnJuego2.textContent = '🎮 Win Juego';

    const btnPartida2 = document.createElement('button');
    btnPartida2.className = 'btn-win-partida';
    btnPartida2.textContent = esBracket ? '🏆 Win Partida' : '✅ Fin Match';

    divBotones2.appendChild(btnJuego2);
    divBotones2.appendChild(btnPartida2);
    j2.appendChild(spanJ2);
    j2.appendChild(divBotones2);

    // Ocultar si hay "Libre" o "Ganador anterior"
    if (jugador1Nombre === 'Libre' || jugador1Nombre === 'Ganador anterior') divBotones1.style.display = 'none';
    if (jugador2Nombre === 'Libre' || jugador2Nombre === 'Ganador anterior') divBotones2.style.display = 'none';

    // Información del Sorteo
    const infoSorteo = document.createElement('span');
    infoSorteo.className = 'info-juego';
    infoSorteo.innerHTML = "🕹️ <strong>Juego:</strong> Pendiente";
    infoSorteo.style.margin = "8px 0";
    infoSorteo.style.textAlign = "center";

    const btnSorteo = document.createElement('button');
    btnSorteo.className = 'btn-sorteo-duelo';
    btnSorteo.textContent = "🎰 Sortear Juego";
    btnSorteo.onclick = () => abrirSorteoDuelo(infoSorteo);

    // Eventos (Win Juego - Historial)
    btnJuego1.onclick = () => registrarGanadorJuego(spanJ1.textContent, divEnfrentamiento);
    btnJuego2.onclick = () => registrarGanadorJuego(spanJ2.textContent, divEnfrentamiento);

    // Eventos (Win Partida)
    if (esBracket) {
        btnPartida1.onclick = () => funcionAvanzarJ1(divEnfrentamiento);
        btnPartida2.onclick = () => funcionAvanzarJ2(divEnfrentamiento);
    } else {
        // En Fase de Grupos, solo bloquea el duelo y registra el final en el historial
        btnPartida1.onclick = () => finalizarMatchGrupo(spanJ1.textContent, divEnfrentamiento);
        btnPartida2.onclick = () => finalizarMatchGrupo(spanJ2.textContent, divEnfrentamiento);
    }

    divEnfrentamiento.appendChild(j1);
    divEnfrentamiento.appendChild(j2);
    divEnfrentamiento.appendChild(infoSorteo);
    divEnfrentamiento.appendChild(btnSorteo);

    return divEnfrentamiento;
}

// Funciones de Historial compartidas
function registrarGanadorJuego(nombre, divEnfrentamientoAct) {
    if (nombre === 'Ganador anterior' || nombre === 'Libre') return;
    const infoSorteoAct = divEnfrentamientoAct.querySelector('.info-juego');
    let juegoJugado = "No sorteado (Por defecto)";
    if (infoSorteoAct) {
        let txt = infoSorteoAct.textContent.replace('🕹️', '').replace('Juego:', '').trim();
        if (txt !== 'Pendiente') juegoJugado = txt;
    }
    const li = document.createElement('li');
    li.style.borderBottom = "1px dashed #ccc";
    li.style.padding = "4px 0";
    li.style.marginLeft = "20px";
    li.innerHTML = `🎮 Juego: <span style="color:#3498db; font-weight:bold;">${juegoJugado}</span> ➔ Ganador: <span style="color:#27ae60; font-weight:bold;">${nombre}</span>`;
    listaHistorico.appendChild(li);
}

function finalizarMatchGrupo(nombre, divEnfrentamientoAct) {
    divEnfrentamientoAct.querySelectorAll('button').forEach(b => b.remove());
    const li = document.createElement('li');
    li.style.borderBottom = "2px solid #bdc3c7";
    li.style.padding = "10px 0";
    li.style.marginTop = "10px";
    li.style.backgroundColor = "#fdfefe";
    li.innerHTML = `✅ <strong>Match de Grupo Resuelto:</strong> 🏆 <strong><span style="color:#f39c12">${nombre}</span></strong> se lleva la partida.`;
    listaHistorico.appendChild(li);
}

// --- 5. GENERACIÓN DEL BRACKET ---
function generarBracket(listaParticipantesFinales) {
    let mezclados = [...listaParticipantesFinales];
    mezclados.sort(() => Math.random() - 0.5);

    let potencia = 2;
    while (potencia < mezclados.length) { potencia *= 2; }

    const faltantes = potencia - mezclados.length;
    for (let i = 0; i < faltantes; i++) { mezclados.push("Libre"); }

    contenedorTorneo.innerHTML = '';
    const divBracket = document.createElement('div');
    divBracket.className = 'bracket';

    let numRondas = Math.log2(potencia);
    let estructuraNodos = [];

    for (let r = 0; r <= numRondas; r++) {
        estructuraNodos[r] = [];
        const divRonda = document.createElement('div');
        divRonda.className = 'ronda';

        if (r === numRondas) {
            const cajaCampeon = document.createElement('div');
            cajaCampeon.className = 'jugador campeon';
            cajaCampeon.innerHTML = `<span class="nombre">🏆 Campeón</span>`;
            estructuraNodos[r].push(cajaCampeon);
            divRonda.appendChild(cajaCampeon);
        } else {
            const numEnfrentamientos = potencia / Math.pow(2, r + 1);

            for (let e = 0; e < numEnfrentamientos; e++) {
                let nombre1 = r === 0 ? mezclados[e * 2] : 'Ganador anterior';
                let nombre2 = r === 0 ? mezclados[e * 2 + 1] : 'Ganador anterior';

                const matchBox = crearCajaMatch(
                    nombre1,
                    nombre2,
                    true,
                    (divAct) => avanzarJugadorBracket(matchBox.querySelectorAll('.nombre')[0].textContent, r, e, divAct, false),
                    (divAct) => avanzarJugadorBracket(matchBox.querySelectorAll('.nombre')[1].textContent, r, e, divAct, false)
                );

                estructuraNodos[r].push({ caja1: matchBox.children[0], caja2: matchBox.children[1] });
                divRonda.appendChild(matchBox);
            }
        }
        divBracket.appendChild(divRonda);
    }
    contenedorTorneo.appendChild(divBracket);

    function avanzarJugadorBracket(nombre, rondaActual, indexEnfrentamiento, divEnfrentamientoAct, esAutoLibre) {
        if (nombre === 'Ganador anterior' || nombre === 'Libre') return;

        if (!esAutoLibre && divEnfrentamientoAct) {
            divEnfrentamientoAct.querySelectorAll('button').forEach(b => b.remove());
            const li = document.createElement('li');
            li.style.borderBottom = "2px solid #bdc3c7";
            li.style.padding = "10px 0";
            li.style.marginTop = "10px";
            li.style.backgroundColor = "#fdfefe";
            li.innerHTML = `🚀 <strong>Match Eliminatorio ${contadorPartidas} Resuelto:</strong> 🏆 <strong><span style="color:#f39c12">${nombre}</span></strong> avanza a la siguiente ronda.`;
            listaHistorico.appendChild(li);
            contadorPartidas++;
        }

        const nextRonda = rondaActual + 1;
        const nextIndex = Math.floor(indexEnfrentamiento / 2);
        const vaArriba = (indexEnfrentamiento % 2 === 0);

        if (nextRonda === numRondas) {
            const cajaCamp = estructuraNodos[nextRonda][0];
            cajaCamp.querySelector('.nombre').textContent = '🏆 ' + nombre;
            cajaCamp.style.backgroundColor = '#f1c40f';
        } else {
            const siguienteCaja = vaArriba ? estructuraNodos[nextRonda][nextIndex].caja1 : estructuraNodos[nextRonda][nextIndex].caja2;
            siguienteCaja.querySelector('.nombre').textContent = nombre;
            siguienteCaja.style.backgroundColor = '#d5f5e3';

            const sBotones = siguienteCaja.querySelector('.botones-container');
            if (sBotones) sBotones.style.display = 'flex';

            const oponente = vaArriba ? estructuraNodos[nextRonda][nextIndex].caja2 : estructuraNodos[nextRonda][nextIndex].caja1;
            if (oponente.querySelector('.nombre').textContent === 'Libre') {
                setTimeout(() => avanzarJugadorBracket(nombre, nextRonda, nextIndex, siguienteCaja.parentElement, true), 400);
            }
        }
    }

    // Pases automáticos por "Libre"
    setTimeout(() => {
        for (let e = 0; e < potencia / 2; e++) {
            const nom1 = mezclados[e * 2];
            const nom2 = mezclados[e * 2 + 1];
            const parentDiv = estructuraNodos[0][e].caja1.parentElement;
            if (nom2 === 'Libre') avanzarJugadorBracket(nom1, 0, e, parentDiv, true);
            if (nom1 === 'Libre') avanzarJugadorBracket(nom2, 0, e, parentDiv, true);
        }
    }, 500);
}

// --- 6. RULETA LOGICA ---
function abrirSorteoDuelo(elementoTextoInfo) {
    dueloActual = elementoTextoInfo;
    if (juegosDisponibles.length === 0) {
        juegosDisponibles = [...listaJuegos];
        alert("¡Se han jugado todos los juegos! La ruleta se ha reiniciado.");
    }
    juegoAsignado.textContent = "";
    ruletaPantalla.innerHTML = "¡Gira para elegir el juego!";
    btnGirar.disabled = false;
    btnGirar.textContent = "Girar Ruleta";
    btnGirar.style.backgroundColor = '#3498db';
    modalRuleta.classList.remove('oculto');
    rotacionAcumulada = 0;
    canvasRuleta.style.transition = 'none';
    canvasRuleta.style.transform = `rotate(0deg)`;
    dibujarRuleta();
    canvasRuleta.offsetHeight;
    canvasRuleta.style.transition = 'transform 4s cubic-bezier(0.25, 0.1, 0.25, 1)';
}

btnCerrarModal.onclick = () => modalRuleta.classList.add('oculto');

function dibujarRuleta() {
    if (!ctx) return;
    ctx.clearRect(0, 0, canvasRuleta.width, canvasRuleta.height);
    const centroX = canvasRuleta.width / 2;
    const centroY = canvasRuleta.height / 2;
    const numSlices = juegosDisponibles.length;
    const anguloPorSlice = (2 * Math.PI) / numSlices;

    for (let i = 0; i < numSlices; i++) {
        const aI = i * anguloPorSlice, aF = aI + anguloPorSlice;
        ctx.beginPath();
        ctx.fillStyle = coloresRuleta[i % coloresRuleta.length];
        ctx.moveTo(centroX, centroY);
        ctx.arc(centroX, centroY, centroX, aI, aF);
        ctx.fill(); ctx.stroke();
        ctx.save();
        ctx.translate(centroX, centroY);
        ctx.rotate(aI + anguloPorSlice / 2);
        ctx.textAlign = "right"; ctx.fillStyle = "#fff"; ctx.font = "bold 16px Arial";
        let txt = juegosDisponibles[i];
        if (txt.length > 15) txt = txt.substring(0, 15) + '...';
        ctx.fillText(txt, centroX - 20, 5);
        ctx.restore();
    }
}

let indiceGanadorActual = -1;
btnGirar.addEventListener('click', () => {
    if (juegosDisponibles.length === 0) return;
    btnGirar.disabled = true;
    ruletaPantalla.textContent = 'Girando...';
    const numSlices = juegosDisponibles.length;
    indiceGanadorActual = Math.floor(Math.random() * numSlices);
    const gradosPorSlice = 360 / numSlices;
    const posObj = 360 - ((indiceGanadorActual * gradosPorSlice) + (gradosPorSlice / 2));
    rotacionAcumulada += (360 * 5) + posObj - (rotacionAcumulada % 360);
    canvasRuleta.style.transform = `rotate(${rotacionAcumulada}deg)`;
});

canvasRuleta.addEventListener('transitionend', () => {
    if (indiceGanadorActual === -1) return;
    const juegoGanador = juegosDisponibles[indiceGanadorActual];
    juegosDisponibles.splice(indiceGanadorActual, 1);
    ruletaPantalla.innerHTML = `🎉 <strong>${juegoGanador}</strong> 🎉`;
    btnGirar.textContent = "Juego Seleccionado";
    btnGirar.style.backgroundColor = '#27ae60';
    if (dueloActual) dueloActual.innerHTML = `🕹️ <strong>Juego:</strong> <span style="color:#e74c3c">${juegoGanador}</span>`;
    setTimeout(() => { modalRuleta.classList.add('oculto'); indiceGanadorActual = -1; }, 2000);
});