// CCT3.js — Parse Excel and create printable QR cards

const fileInput = document.getElementById('fileInput');
const cardsContainer = document.getElementById('cardsContainer');
const printBtn = document.getElementById('printBtn');


fileInput.addEventListener('change', handleFile, false);
printBtn.addEventListener('click', () => window.print());


function handleFile(e){
const f = e.target.files[0];
if(!f) return;
const reader = new FileReader();
reader.onload = function(evt){
const data = evt.target.result;
let workbook;
try{
workbook = XLSX.read(data, {type: 'array'});
}catch(err){
alert('No se pudo leer el archivo Excel. Asegúrate de que es .xlsx o .xls');
console.error(err);
return;
}
const firstSheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[firstSheetName];
const rows = XLSX.utils.sheet_to_json(worksheet, {defval: ''});
renderCards(rows);
};
reader.readAsArrayBuffer(f);
}

// Intenta emparejar claves tolerando diferencias de espacios/puntos/minúsculas
function findKey(keys, want){
const normalize = s => String(s).replace(/\s+/g, ' ').trim().toLowerCase().replace(/\./g,'');
const target = normalize(want);
for(const k of keys){
if(normalize(k) === target) return k;
}
// fallback: contiene palabras clave
for(const k of keys){
const nk = normalize(k);
const parts = target.split(' ');
if(parts.every(p => nk.includes(p))) return k;
}
return null;
}

function renderCards(rows){
cardsContainer.innerHTML = '';
if(!rows || rows.length === 0){
cardsContainer.textContent = 'No hay filas en la hoja.';
return;
}


// Obtener claves del primer objeto
const keys = Object.keys(rows[0]);


const mapping = {
userField: findKey(keys, 'Campo definido por el usuario 1'),
urbantzId: findKey(keys, 'Id. carga de Urbantz') || findKey(keys, 'Id. carga Urbantz') || findKey(keys, 'Id carga de Urbantz'),
cargaId: findKey(keys, 'Id. de la carga') || findKey(keys, 'Id de la carga') || findKey(keys, 'Id carga'),
almacen: findKey(keys, 'Almacén') || findKey(keys, 'Almacen'),
transportista: findKey(keys, 'Transportista de envío') || findKey(keys, 'Transportista'),
envios: findKey(keys, 'Envíos') || findKey(keys, 'Envios')
};


rows.forEach((row, idx) => {
const card = document.createElement('article');
card.className = 'card';


// QR
const qrArea = document.createElement('div');
qrArea.className = 'qr-area';
const qrDiv = document.createElement('div');
qrDiv.className = 'qr';
qrDiv.id = `qr-${idx}`;
qrArea.appendChild(qrDiv);


// Fields
const fields = document.createElement('div');
fields.className = 'fields';


function makeRow(labelText, valueText){
const rowEl = document.createElement('div');
rowEl.className = 'field-row';
const lbl = document.createElement('div'); lbl.className = 'label'; lbl.textContent = labelText;
const val = document.createElement('div'); val.className = 'value'; val.textContent = valueText;
rowEl.appendChild(lbl); rowEl.appendChild(val);
return rowEl;
}


const userVal = mapping.userField ? row[mapping.userField] : '';
const urbVal = mapping.urbantzId ? row[mapping.urbantzId] : '';
const cargaVal = mapping.cargaId ? row[mapping.cargaId] : '';
const almacVal = mapping.almacen ? row[mapping.almacen] : '';
const transpVal = mapping.transportista ? row[mapping.transportista] : '';
const envValRaw = mapping.envios ? row[mapping.envios] : '';
const envVal = envValRaw === '' ? '' : `envíos ${envValRaw}`;


fields.appendChild(makeRow('Campo definido:', userVal));
fields.appendChild(makeRow('Id. carga Urbantz:', urbVal));
fields.appendChild(makeRow('Id. de la carga:', cargaVal));
fields.appendChild(makeRow('Almacén:', almacVal));
fields.appendChild(makeRow('Transportista:', transpVal));
fields.appendChild(makeRow('Envíos:', envVal));


card.appendChild(qrArea);
card.appendChild(fields);


cardsContainer.appendChild(card);


// Generar QR con el valor de la columna "Id. de la carga"
const qrText = cargaVal !== undefined && String(cargaVal).trim() !== '' ? String(cargaVal) : '';
// Si no hay valor, generamos un QR vacío / o con texto indicando falta
const qrcode = new QRCode(qrDiv, {
text: qrText || ' ',
width: 200,
height: 200,
correctLevel: QRCode.CorrectLevel.M
});
});
}