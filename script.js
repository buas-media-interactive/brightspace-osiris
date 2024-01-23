const pointsrx = / Points Grade <[^>]+MaxPoints:(\d+)/;
const numeratorx = / Numerator$/;
const studentidrx = /\d{6}/;
let brightspace_content = [];
let selected_grades = [];
let name_columns = [0];
let grade_columns = [];

function selectGrades() {
    let column = parseInt(document.getElementById('gradecolumn').value);
    if (!isNaN(column)) {
        let filled_in_grades = brightspace_content.slice(1).filter(
            c => (c[column] !== undefined)
        );
        let full_column = grade_columns[column];
        selected_grades = filled_in_grades.map(g => {
            let denominator = full_column.denominator || g[column+1];
            return [
                studentidrx.exec(g[0])[0],
                name_columns.map(n => g[n]).join(', '),
                Math.round(100*g[column]/denominator)/10,
                `${g[column]}/${denominator}`
            ]
        });
        document.getElementById('gradelist').innerHTML = selected_grades.map(g => {
            let cls = (g[2] < 5.5) ? 'table-danger' : 'table-success';
            let line = g.map(c => `<td>${c}</td>`).join('');
            return `<tr class="${cls}">${line}</tr>`;
        }
        ).join('');
    } else {
        selected_grades = [];
        console.log('No valid grade column');
    }
}

document.getElementById('bspfile').addEventListener('change', e => {
    let file = e.target.files[0];
    if (!file) {
        return;
    }
    var reader = new FileReader();
    reader.onload = f => {
        let workbook = XLSX.read(reader.result);
        let worksheet = workbook.Sheets[workbook.SheetNames[0]];
        brightspace_content = XLSX.utils.sheet_to_json(worksheet, {header: 1});
        name_columns = [];
        grade_columns = brightspace_content[0].map((k, n) => {
            if (/Name$/.test(k)) {
                name_columns.push(n);
            }
            if (pointsrx.test(k)) {
                return {
                    'column': n,
                    'full_name': k,
                    'short_name': k.split(pointsrx)[0],
                    'denominator': parseInt(k.split(pointsrx)[1])
                }
            }
            if (numeratorx.test(k)) {
                return {
                    'column': n,
                    'full_name': k,
                    'short_name': k.split(numeratorx)[0],
                    'denominator': null
                }
            }
            return null;
        });
        if (!name_columns) {
            name_columns = [0];
        }
        document.getElementById('gradecolumn').innerHTML = grade_columns.filter(gk => gk).map(gk => {
            return `<option value="${gk.column}">${gk.short_name}</option>`;
        }).join('');
        selectGrades();
    };
    reader.readAsArrayBuffer(file);
});

document.getElementById('gradecolumn').addEventListener('change', selectGrades);