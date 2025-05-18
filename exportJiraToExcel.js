const axios = require('axios');
const ExcelJS = require('exceljs');
require('dotenv').config();
const dayjs = require('dayjs');
const localeEs = require('dayjs/locale/es');
const isSameOrBefore = require('dayjs/plugin/isSameOrBefore');
const advancedFormat = require('dayjs/plugin/advancedFormat');
dayjs.locale('es');
dayjs.extend(isSameOrBefore);
dayjs.extend(advancedFormat);

const JIRA_DOMAIN = 'https://desarrollosica.atlassian.net';
const EMAIL = 'jhonalejoo@gmail.com';
const NAME = 'Jhon Alejandro Cuervo Sanchez';
const API_TOKEN = process.env.API_TOKEN;

const hoy = dayjs();
const primerDia = hoy.startOf('month').format('YYYY-MM-DD');
const ultimoDia = hoy.endOf('month').format('YYYY-MM-DD');
const periodoTexto = `${dayjs(primerDia).format('DD/MM/YYYY')}-${dayjs(ultimoDia).format('DD/MM/YYYY')}`;
const JQL = `project = "ScrumSica" AND assignee = currentUser() AND duedate >= ${primerDia} AND duedate <= ${ultimoDia}`;

// 👇 Cargar festivos desde la API pública
async function obtenerFestivosColombia(year) {
  try {
    const response = await axios.get(`https://date.nager.at/api/v3/PublicHolidays/${year}/CO`);
    return response.data.map(f => f.date); // ["2025-01-01", ...]
  } catch (err) {
    console.warn("⚠️ No se pudieron cargar los festivos:", err.message);
    return [];
  }
}

async function exportJiraToExcel() {
  try {
    const festivos = await obtenerFestivosColombia(hoy.year());

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

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Reporte');

    sheet.mergeCells('A1:H1');
    const tituloCell = sheet.getCell('A1');
    tituloCell.value = 'REPORTE ACTIVIDADES FEDERECAFE';
    tituloCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF002060' } };
    tituloCell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 14 };
    tituloCell.alignment = { vertical: 'middle', horizontal: 'center' };

    sheet.getRow(2).values = ['TRABAJADOR', NAME, `Periodo: ${periodoTexto}`];
    sheet.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF002060' } };
    sheet.getCell('A2').font = { color: { argb: 'FFFFFFFF' }, bold: true };
    sheet.getCell('B2').font = { color: { argb: 'FF002060' }, bold: true };
    sheet.getCell('C2').font = { color: { argb: 'FF002060' }, bold: true };

    const encabezados = ['FECHA', 'ID CASO', 'FECHA REGISTRO', 'ASUNTO', 'CATEGORIA', 'GRUPO DE ESPECIALISTAS', 'ESTADO', 'FECHA DE SOLUCIÓN'];
    sheet.addRow(encabezados);
    const headerRow = sheet.getRow(3);
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF002060' } };
      cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
    });

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
        const diaSemana = d.day(); // 0 = domingo, 6 = sábado
        const dStr = d.format('YYYY-MM-DD');
        const dExcel = d.format('DD/MM/YYYY');

        const esFestivo = festivos.includes(dStr);
        const esFinDeSemana = diaSemana === 0 || diaSemana === 6;

        if (esFestivo || esFinDeSemana) {
          const textoDia = esFestivo ? 'FESTIVO' : diaSemana === 0 ? 'DOMINGO' : 'SÁBADO';
          const row = sheet.addRow([
            dExcel,
            '', '', textoDia, '', '', '', ''
          ]);
          row.eachCell((cell, colNumber) => {
            if (colNumber >= 2 && colNumber <= 8) {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD9D9D9' }
              };
              cell.font = { italic: true };
              cell.alignment = { vertical: 'middle', horizontal: 'center' };
            }
          });
        } else {
          const esUltimoDia = d.isSame(fechaFin, 'day');
          sheet.addRow([
            dExcel,
            idCaso,
            fechaInicio.format('DD/MM/YYYY'),
            asunto,
            'Web/movil',
            'Desarrolladores',
            esUltimoDia ? 'Cerrado' : 'Abierto',
            esUltimoDia ? fechaFin.format('DD/MM/YYYY') : ''
          ]);
        }

        d = d.add(1, 'day');
      }
    }

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

    sheet.columns.forEach((col) => {
      col.width = 25;
    });

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
