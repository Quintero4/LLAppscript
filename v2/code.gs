function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Lessons Learnt App')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function guardarDatos(data, fileData) {
  try {
    const SHEET_ID = '1C-tLy2n9-W5RVgdcF6P6BNmkBhlBY8ePstOSCFJKNF0'; 
    const FOLDER_ID = '1_1rdyGfWbCGxIoOPQh9itc4bz7mo2HZY'; 
    const TEMPLATE_ID = '1IUSitAQPApYRZO1tXRpT0UyOM2NE5rCsQISmVtIKyIs'; // <--- PEGA AQUÍ EL ID DE TU SLIDE PLANTILLA

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheets()[0];
    const folder = DriveApp.getFolderById(FOLDER_ID);

    // 1. CREAR COPIA DE LA PLANTILLA
    const nombrePPT = "Lesson_Learnt_" + (data.title || "Sin_Titulo") + "_" + new Date().getTime();
    const copiaArchivo = DriveApp.getFileById(TEMPLATE_ID).makeCopy(nombrePPT, folder);
    const presentation = SlidesApp.openById(copiaArchivo.getId());
    const slide = presentation.getSlides()[0];

    // 2. REEMPLAZAR TEXTOS (Busca {{etiqueta}} y pone el valor)
    slide.replaceAllText('{{title}}', data.title || "");
    slide.replaceAllText('{{sta}}', data.sta || "");
    slide.replaceAllText('{{msn}}', data.msn || "");
    slide.replaceAllText('{{author}}', data.author || "");
    slide.replaceAllText('{{Reference}}', data.Reference || "");
    slide.replaceAllText('{{description}}', data.description || "");
    slide.replaceAllText('{{solution}}', data.solution || "");
    slide.replaceAllText('{{do}}', data.do || "");
    slide.replaceAllText('{{dont}}', data.dont || "");
    slide.replaceAllText('{{impactedProcess}}', data.impactedProcess || "");
    slide.replaceAllText('{{actions}}', data.actions || "");
    slide.replaceAllText('{{date}}', new Date().toLocaleDateString());


    // 3. INSERTAR IMAGEN
    if (fileData && fileData.data) {
      const bytes = Utilities.base64Decode(fileData.data);
      const blob = Utilities.newBlob(bytes, fileData.type, fileData.name);
      
//Buscamos la forma
const shapes = slide.getShapes();
let marcadorEncontrado = false;

for (let i=0; i < shapes.length; i++){
  const shape = shapes [i];
  const text = shape.getText().asString().trim();

  //Si la forma contiene nuestra etiqueta de la imagen
  if (text === '{{Image}}') {
    //obtenemos posición y dimensiones de la forma
    const left = shape.getLeft();
    const top = shape.getTop();
    const width = shape.getWidth();
    const height = shape.getHeight();

    //Insertamos la imagen escalada al tamaño de la forma
    slide.insertImage(blob, left, top, width,height);

    //borramos la forma original para que no se vea el texto  {{Image}}
    shape.remove();
    marcadorEncontrado = true;
    break;
  }
}

//Si no encontramos el marcador, podemos inserta por defecto en una esquina
if (!marcadorEncontrado) {
  console.warn ("No se encontró el marcador {{Image}} en la plantilla");
}
    }

    presentation.saveAndClose();
    const finalUrl = copiaArchivo.getUrl();

    // 4. REGISTRAR EN EXCEL
    const row = [
      data.sta || "", data.msn || "", data.author || "", data.Reference || "",
      new Date(), data.title || "", data.description || "", data.solution || "",
      data.do || "", data.dont || "", data.impactedProcess || "", data.actions || "",
      finalUrl
    ];

    sheet.appendRow(row);
    return "✅ Guardado. Se ha creado una ficha idéntica a la plantilla.";

  } catch (e) {
    throw new Error("Error: " + e.message);
  }
}

