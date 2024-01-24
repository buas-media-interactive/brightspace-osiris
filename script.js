const pointsrx = / Points Grade <[^>]+MaxPoints:(\d+)/;
const numeratorx = / Numerator$/;
const studentidrx = /\d{6}/;
let brightspace_content = [];
let selected_grades = [];
let name_columns = [0];
let grade_columns = [];
let now = new Date();
let ac_year = now.getFullYear();
if (now.getMonth() < 8) {
    ac_year--;
}
let now_date = now.toISOString().split(/[-T]/).slice(0,3).reverse().join('-');

let osiris_content = [
    ["Course","COURSE-CODE",undefined,undefined,"Time"],
    ["Name","Course Name"],
    ["Academic year",ac_year.toString()],
    ["Test","TESTCODE","Assignment Name"],
    ["Block","YEAR",undefined,"Grading scale","1-DEC NUM+GK"],
    ["Opportunity","1"],
    [],
    ["Student number","Name","Test date","Grade"]
];

function getIntersection() {
    let osiris_student_ids = osiris_content.slice(8).map(c => c[0]);
    return selected_grades.filter(c => osiris_student_ids.includes(c[0]));
}

function displayIntersection() {
    let intersectiondescription = 'No grades to export';
    if (selected_grades.length) {
        if (osiris_content.length < 9) {
            intersectiondescription = 'The Osiris file is blank. All grades from Brightspace will be exported';
        } else {
            intersectiondescription = getIntersection().length + ' students are both in the Osiris and Brightspace lists. Only these students will be exported';
        }
    }
    document.getElementById('contentdescription').textContent = intersectiondescription;
}

function generateSpreadsheet() {
    let grades_to_export = (osiris_content.length < 9) ? selected_grades : getIntersection();
    if (!grades_to_export.length) {
        alert('No grades to export');
    }
    let final_table = osiris_content.slice(0,8).concat(
        grades_to_export.map(g => [
            g[0],
            g[1],
            now_date,
            g[2].toString().replace('.',','),
        ])
    );
    let workbook = XLSX.utils.book_new();
    let worksheet = XLSX.utils.aoa_to_sheet(final_table);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Test list");
    XLSX.writeFile(
        workbook,
        `Generated-Osiris-${osiris_content[0][1]}-${now_date}.xlsx`,
        { compression: true, ignoreEC: false, bookSST: true }
    );
}

function displayOsiris() {
    document.getElementById('osirisheaders').innerHTML = osiris_content.slice(0,6).map(
        l => '<tr>' + l.map(c => `<td>${c||''}</td>`).join('') + '</tr>'
    ).join('');
    document.getElementById('osiriscount').textContent = osiris_content.slice(8).length;
    displayIntersection();
}

function selectGrades() {
    let column = parseInt(document.getElementById('gradecolumn').value);
    if (isNaN(column)) {
        selected_grades = [];
        document.getElementById('bspcolname').textContent = "(no grade selected)";
    } else {
        let filled_in_grades = brightspace_content.slice(1).filter(
            c => (c[column] !== undefined)
        );
        let full_column = grade_columns[column];
        selected_grades = filled_in_grades.map(g => {
            let denominator = full_column.denominator || g[column+1];
            return [
                (studentidrx.exec(g[0])||["Invalid Student ID"])[0],
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
        document.getElementById('bspcolname').textContent = full_column.full_name;
    }
    document.getElementById('bspcount').textContent = selected_grades.length;
    displayIntersection();
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
        if (!name_columns.length) {
            name_columns = [0];
        }
        document.getElementById('gradecolumn').innerHTML = grade_columns.filter(gk => gk).map(gk => {
            return `<option value="${gk.column}">${gk.short_name}</option>`;
        }).join('');
        selectGrades();
    };
    reader.readAsArrayBuffer(file);
});

document.getElementById('osirisfile').addEventListener('change', e => {
    let file = e.target.files[0];
    if (!file) {
        return;
    }
    var reader = new FileReader();
    reader.onload = f => {
        let workbook = XLSX.read(reader.result);
        let worksheet = workbook.Sheets[workbook.SheetNames[0]];
        osiris_content = XLSX.utils.sheet_to_json(worksheet, {header: 1});
        displayOsiris();
    };
    reader.readAsArrayBuffer(file);
});

document.getElementById('gradecolumn').addEventListener('change', selectGrades);
document.getElementById('exportbutton').addEventListener('click', generateSpreadsheet);

displayOsiris();
