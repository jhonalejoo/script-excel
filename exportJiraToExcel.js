const axios = require('axios');
const ExcelJS = require('exceljs');
require('dotenv').config(); // Esto carga las variables del archivo .env
const dayjs = require('dayjs');
const localeEs = require('dayjs/locale/es'); // <- importa la localización en español
dayjs.locale('es');
const isSameOrBefore = require('dayjs/plugin/isSameOrBefore');
const advancedFormat = require('dayjs/plugin/advancedFormat');
dayjs.extend(isSameOrBefore);
dayjs.extend(advancedFormat);

// Configura tus credenciales
const JIRA_DOMAIN = 'https://desarrollosica.atlassian.net';
const EMAIL = 'jhonalejoo@gmail.com';
const NAME = 'Jhon Alejandro Cuervo Sanchez';
const API_TOKEN = process.env.ATLASSIAN_TOKEN;

// Genera fechas del mes actual
const hoy = dayjs();
const primerDia = hoy.startOf('month').format('YYYY-MM-DD');
const ultimoDia = hoy.endOf('month').format('YYYY-MM-DD');
const periodoTexto = `${dayjs(primerDia).format('DD/MM/YYYY')}-${dayjs(ultimoDia).format('DD/MM/YYYY')}`;

console.log('Fechas del mes actual:', primerDia, 'a', ultimoDia);
// JQL: Tareas del usuario en el mes actual
const JQL = `project = "ScrumSica" AND assignee = currentUser() AND duedate >= "2025-05-01" AND duedate <= "2025-05-31"`;

// Función principal
async function exportJiraToExcel() {
  try {
    const response = await axios.get(`${JIRA_DOMAIN}/rest/api/3/search`, {
      headers: { 'Accept': 'application/json' },
      auth: { username: EMAIL, password: API_TOKEN },
      params: {
        jql: JQL,
        maxResults: 100,
        fields: 'summary,status,assignee,created,updated,duedate,customfield_10015,customfield_10034'
      }
    });

    const issues = response.data.issues;

    issues.sort((a, b) => new Date(a.fields.duedate) - new Date(b.fields.duedate));

    // Crear el Excel
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Reporte');

    // Primera fila (título grande)
    sheet.mergeCells('A1:H1');
    const tituloCell = sheet.getCell('A1');
    tituloCell.value = 'REPORTE ACTIVIDADES FEDERECAFE';
    tituloCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF002060' }
    };
    tituloCell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 14 };
    tituloCell.alignment = { vertical: 'middle', horizontal: 'center' };

    // Segunda fila: TRABAJADOR | Nombre | Periodo
    sheet.getRow(2).values = ['TRABAJADOR', NAME, `Periodo: ${periodoTexto}`];
    sheet.getCell('A2').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF002060' }
    };
    sheet.getCell('A2').font = { color: { argb: 'FFFFFFFF' }, bold: true };
    sheet.getCell('B2').font = { color: { argb: 'FF002060' }, bold: true };
    sheet.getCell('C2').font = { color: { argb: 'FF002060' }, bold: true };

    // Tercera fila: Encabezados
    const encabezados = ['FECHA', 'ID CASO', 'FECHA REGISTRO', 'ASUNTO', 'CATEGORIA', 'GRUPO DE ESPECIALISTAS', 'ESTADO', 'FECHA DE SOLUCIÓN'];
    sheet.addRow(encabezados);
    const headerRow = sheet.getRow(3);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF002060' }
      };
      cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
    });

    // Agregar datos
    for (const issue of issues) {
      const fechaRegistro = issue.fields['customfield_10015'];
      const fechaFinCruda = issue.fields['duedate'];

      if (!fechaRegistro || !fechaFinCruda) continue;

      const fechaInicio = dayjs(fechaRegistro);
      const fechaFin = dayjs(fechaFinCruda);

      if (!fechaInicio.isValid() || !fechaFin.isValid()) continue;

      const idCaso = issue.fields['customfield_10034'] || '';
      const asunto = issue.fields.summary || '';

      let d = fechaInicio.clone();
      while (d.isSame(fechaFin) || d.isBefore(fechaFin)) {
        const esUltimoDia = d.isSame(fechaFin, 'day');

        sheet.addRow([
          d.format('DD/MM/YYYY'),
          idCaso,
          fechaInicio.format('DD/MM/YYYY'),
          asunto,
          'Web/movil',
          'Desarrolladores',
          esUltimoDia ? 'Cerrado' : 'Abierto',
          esUltimoDia ? fechaFin.format('DD/MM/YYYY') : ''
        ]);

        d = d.add(1, 'day');
      }
    }

      // Aplicar bordes y auto ajuste de texto
      const aplicarEstilos = (fila) => {
        fila.eachCell({ includeEmpty: true }, (cell) => {
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF002060' } },
            left: { style: 'thin', color: { argb: 'FF002060' } },
            bottom: { style: 'thin', color: { argb: 'FF002060' } },
            right: { style: 'thin', color: { argb: 'FF002060' } }
          };
          cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
        });
      };
  
      for (let i = 3; i <= sheet.rowCount; i++) {
        aplicarEstilos(sheet.getRow(i));
      }
  
      // Ajustar ancho de columnas
      sheet.columns.forEach((col) => {
        col.width = 25;
      });

    // Nombre del archivo dinámico
    const mes = hoy.format('MMMM');
    const mesCapitalizado = mes.charAt(0).toUpperCase() + mes.slice(1);
    const año = hoy.format('YYYY');
    const nombreArchivo = `Reporte Samtel_FEDERACAFE_${NAME}_${mesCapitalizado} ${año}.xlsx`;

    await workbook.xlsx.writeFile(nombreArchivo);
    console.log(`✅ Archivo Excel generado correctamente: ${nombreArchivo}`);
  } catch (error) {
    console.error('❌ Error al consultar Jira o generar Excel:', error.response?.data || error.message);
  }
}

exportJiraToExcel();
