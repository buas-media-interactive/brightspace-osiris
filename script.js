const pointsrx = / Points Grade <[^>]+MaxPoints:(\d+)/;
const numeratorx = / Numerator$/;
const studentidrx = /\d{6}/;
const osiris_fake_headers = [
    [],
    ["PLEASE COPY THE CONTENTS BELOW"],
    [],
    [],
    ["INTO A VALID OSIRIS SPREADSHEET"],
    [],
    [],
    ["Student number","Name","Test date","Grade"],
];

let osiris_headers = [];

let course_code = 'manual-merge';
let brightspace_content = [];
let selected_grades = [];
let name_columns = [0];
let grade_columns = [];
let now = new Date();
let now_date = now.toISOString().split(/[-T]/).slice(0,3).reverse().join('-');

let osiris_students_ids = [];
let using_real_osiris_spreadsheet = false;
let osiris_workbook = XLSX.utils.book_new();
let osiris_worksheet = XLSX.utils.aoa_to_sheet(osiris_fake_headers);
XLSX.utils.book_append_sheet(osiris_workbook, osiris_worksheet, "Test list");

const header_range_end = 7;

function deleteAllGrades() {
    let range = XLSX.utils.decode_range(osiris_worksheet["!ref"]);
    console.log(JSON.stringify(range));
    if (range.e.r > header_range_end) {
        for (let r = header_range_end + 1; r <= range.e.r; r++) {
            for (let c = range.s.c; c <= range.e.c; c++) {
                let addr = XLSX.utils.encode_cell({r:r, c:c});
                delete osiris_worksheet[addr];
            }
        }
        range.e.r = header_range_end;
        osiris_worksheet["!ref"] = XLSX.utils.encode_range(range);
    }   
}

function getIntersection() {
    return selected_grades.filter(c => osiris_students_ids.includes(c[0]));
}

function displayIntersection() {
    let intersectiondescription = 'No grades to export';
    if (selected_grades.length) {
        if (!using_real_osiris_spreadsheet) {
            intersectiondescription = 'All grades from Brightspace will be exported in a way that you need to copy and paste manually into an Osiris spreadsheet';
        } else {
            let intersection_length = getIntersection().length;
            intersectiondescription = intersection_length + ' students are both in the Osiris and Brightspace lists.';
            if (intersection_length < selected_grades.length) {
                intersectiondescription += ' Only these students will be exported.';
            } 
        }
    }
    document.getElementById('contentdescription').textContent = intersectiondescription;
}

function generateSpreadsheet() {
    let grades_to_export = using_real_osiris_spreadsheet ? getIntersection() : selected_grades;
    if (!grades_to_export.length) {
        alert('No grades to export');
    }
    deleteAllGrades();
    XLSX.utils.sheet_add_aoa(
        osiris_worksheet,
        grades_to_export.map(g => [
            g[0],
            g[1],
            now_date,
            g[2].toString().replace('.',','),
        ]),
        { origin: "A9" }
    );

    XLSX.writeFile(osiris_workbook, `Generated-Osiris-${course_code}-${now_date}.xlsx`,{ compression: true });
}

function displayOsiris() {
    if (using_real_osiris_spreadsheet) {
        document.getElementById('osirisheaders').innerHTML = osiris_headers.map(
            l => '<tr>' + l.map(c => `<td>${c||''}</td>`).join('') + '</tr>'
        ).join('');
        document.getElementById('osiriscount').textContent = osiris_students_ids.length;
        document.getElementById('osirisdata').style.display = 'block';
    } else {
        document.getElementById('osirisdata').style.display = 'none';
    }
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
            let denominator = full_column.denominator || g[column+1] || 100;
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
        osiris_workbook = XLSX.read(reader.result);
        osiris_worksheet = osiris_workbook.Sheets[osiris_workbook.SheetNames[0]];
        using_real_osiris_spreadsheet = true;
        console.log(osiris_worksheet);
        let content = XLSX.utils.sheet_to_json(osiris_worksheet, {header: 1});
        osiris_students_ids = content.slice(8).map(c => c[0]);
        osiris_headers = content.slice(0,6);
        course_code = content[0][1];
        console.log(osiris_students_ids);
        displayOsiris();
    };
    reader.readAsArrayBuffer(file);
});

document.getElementById('gradecolumn').addEventListener('change', selectGrades);
document.getElementById('exportbutton').addEventListener('click', generateSpreadsheet);

