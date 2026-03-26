// --- VARIABLES GLOBALES ---
let listaParticipantes = [];
let listaJuegos = [];
let formatoTorneo = '';
let dueloActual = null;
let rotacionAcumulada = 0;
let juegosDisponibles = [];

// Colores para la ruleta
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

const btnIniciar = document.getElementById('btn-iniciar');
const seccionConfig = document.getElementById('configuracion');
const seccionTorneo = document.getElementById('seccion-torneo');
const contenedorTorneo = document.getElementById('contenedor-torneo');

// Elementos de la Ruleta (Modal)
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


// --- 3. INICIAR EL TORNEO ---
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

    // Inicializamos los juegos disponibles al empezar el torneo
    juegosDisponibles = [...listaJuegos];

    formatoTorneo = listaParticipantes.length <= 5 ? 'Grupos' : 'Bracket';
    document.getElementById('titulo-formato').textContent = `Formato del Torneo: ${formatoTorneo} (${listaParticipantes.length} Jugadores)`;

    seccionConfig.classList.add('oculto');
    seccionTorneo.classList.remove('oculto');

    generarBracket();
});


// --- 4. LA RULETA DEL MODAL ---
function abrirSorteoDuelo(elementoTextoInfo) {
    dueloActual = elementoTextoInfo;

    // Si la lista se quedó vacía, la recargamos para que el torneo pueda continuar
    if (juegosDisponibles.length === 0) {
        juegosDisponibles = [...listaJuegos];
        alert("¡Se han jugado todos los juegos! La ruleta se ha reiniciado para los siguientes duelos.");
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
    const radio = centroX;
    const numSlices = juegosDisponibles.length;
    const anguloPorSlice = (2 * Math.PI) / numSlices;

    for (let i = 0; i < numSlices; i++) {
        const anguloInicio = i * anguloPorSlice;
        const anguloFin = anguloInicio + anguloPorSlice;

        ctx.beginPath();
        ctx.fillStyle = coloresRuleta[i % coloresRuleta.length];
        ctx.moveTo(centroX, centroY);
        ctx.arc(centroX, centroY, radio, anguloInicio, anguloFin);
        ctx.fill();
        ctx.stroke();

        ctx.save();
        ctx.translate(centroX, centroY);
        ctx.rotate(anguloInicio + anguloPorSlice / 2);
        ctx.textAlign = "right";
        ctx.fillStyle = "#fff";
        ctx.font = "bold 16px Arial";

        let texto = juegosDisponibles[i];
        if (texto.length > 15) texto = texto.substring(0, 15) + '...';

        ctx.fillText(texto, radio - 20, 5);
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
    const gradosCentroGanador = (indiceGanadorActual * gradosPorSlice) + (gradosPorSlice / 2);

    const gradosGiroExtra = 360 * 5;
    const posicionObjetivo = 360 - gradosCentroGanador;

    rotacionAcumulada += gradosGiroExtra + posicionObjetivo - (rotacionAcumulada % 360);
    canvasRuleta.style.transform = `rotate(${rotacionAcumulada}deg)`;
});

canvasRuleta.addEventListener('transitionend', () => {
    if (indiceGanadorActual === -1) return;

    const juegoGanador = juegosDisponibles[indiceGanadorActual];

    // AQUÍ ESTÁ LA MAGIA: Eliminamos el juego ganado de la lista disponible
    juegosDisponibles.splice(indiceGanadorActual, 1);

    ruletaPantalla.innerHTML = `🎉 <strong>${juegoGanador}</strong> 🎉`;
    btnGirar.textContent = "Juego Seleccionado";
    btnGirar.style.backgroundColor = '#27ae60';

    if (dueloActual) {
        dueloActual.innerHTML = `🕹️ <strong>Juego:</strong> <span style="color:#e74c3c">${juegoGanador}</span>`;
    }

    setTimeout(() => {
        modalRuleta.classList.add('oculto');
        indiceGanadorActual = -1;
    }, 2000);
});


// --- 5. GENERACIÓN DEL BRACKET Y LÓGICA DE AVANCE ---
function generarBracket() {
    let mezclados = [...listaParticipantes];
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
            cajaCampeon.textContent = '🏆 Campeón';
            estructuraNodos[r].push(cajaCampeon);
            divRonda.appendChild(cajaCampeon);
        } else {
            const numEnfrentamientos = potencia / Math.pow(2, r + 1);

            for (let e = 0; e < numEnfrentamientos; e++) {
                const divEnfrentamiento = document.createElement('div');
                divEnfrentamiento.className = 'enfrentamiento';

                const j1 = document.createElement('div');
                j1.className = 'jugador';
                const j2 = document.createElement('div');
                j2.className = 'jugador';

                j1.style.cursor = "pointer";
                j1.title = "Haz clic para avanzar a este jugador";
                j2.style.cursor = "pointer";
                j2.title = "Haz clic para avanzar a este jugador";

                const infoSorteo = document.createElement('span');
                infoSorteo.className = 'info-juego';
                infoSorteo.innerHTML = "🕹️ <strong>Juego:</strong> Pendiente";
                infoSorteo.style.margin = "8px 0";
                infoSorteo.style.textAlign = "center";

                const btnSorteo = document.createElement('button');
                btnSorteo.className = 'btn-sorteo-duelo';
                btnSorteo.textContent = "🎰 Sortear Juego";
                btnSorteo.onclick = () => abrirSorteoDuelo(infoSorteo);

                if (r === 0) {
                    j1.textContent = mezclados[e * 2];
                    j2.textContent = mezclados[e * 2 + 1];
                } else {
                    j1.textContent = 'Ganador anterior';
                    j2.textContent = 'Ganador anterior';
                }

                j1.onclick = () => avanzarJugador(j1.textContent, r, e, 0);
                j2.onclick = () => avanzarJugador(j2.textContent, r, e, 1);

                estructuraNodos[r].push({ caja1: j1, caja2: j2 });

                divEnfrentamiento.appendChild(j1);
                divEnfrentamiento.appendChild(j2);
                divEnfrentamiento.appendChild(infoSorteo);
                divEnfrentamiento.appendChild(btnSorteo);
                divRonda.appendChild(divEnfrentamiento);
            }
        }
        divBracket.appendChild(divRonda);
    }

    contenedorTorneo.appendChild(divBracket);

    function avanzarJugador(nombre, rondaActual, indexEnfrentamiento, posJugador) {
        if (nombre === 'Ganador anterior' || nombre === 'Libre') return;

        const nextRonda = rondaActual + 1;
        const nextIndex = Math.floor(indexEnfrentamiento / 2);
        const vaArriba = (indexEnfrentamiento % 2 === 0);

        if (nextRonda === numRondas) {
            estructuraNodos[nextRonda][0].textContent = '🏆 ' + nombre;
            estructuraNodos[nextRonda][0].style.backgroundColor = '#f1c40f';
            estructuraNodos[nextRonda][0].style.color = '#333';
        } else {
            const siguienteCaja = vaArriba ? estructuraNodos[nextRonda][nextIndex].caja1 : estructuraNodos[nextRonda][nextIndex].caja2;
            siguienteCaja.textContent = nombre;
            siguienteCaja.style.backgroundColor = '#d5f5e3';

            const oponente = vaArriba ? estructuraNodos[nextRonda][nextIndex].caja2 : estructuraNodos[nextRonda][nextIndex].caja1;
            if (oponente.textContent === 'Libre') {
                setTimeout(() => {
                    avanzarJugador(nombre, nextRonda, nextIndex, vaArriba ? 0 : 1);
                }, 400);
            }
        }
    }

    setTimeout(() => {
        for (let e = 0; e < potencia / 2; e++) {
            const nom1 = mezclados[e * 2];
            const nom2 = mezclados[e * 2 + 1];
            if (nom2 === 'Libre') avanzarJugador(nom1, 0, e, 0);
            if (nom1 === 'Libre') avanzarJugador(nom2, 0, e, 1);
        }
    }, 500);
}