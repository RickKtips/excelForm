// Tell TypeScript that the XLSX variable is available in the global scope
declare var XLSX: any;

document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('excel-form') as HTMLFormElement;
    const addTimeButton = document.getElementById('add-time') as HTMLButtonElement;
    const timeFieldsContainer = document.getElementById('time-fields') as HTMLDivElement;
    const fileInput = document.getElementById('excel-file') as HTMLInputElement;

    addTimeButton.addEventListener('click', () => {
        const timeInputContainer = document.createElement('div');
        timeInputContainer.classList.add('time-input');

        const newTimeInput = document.createElement('input');
        newTimeInput.type = 'time';
        newTimeInput.name = 'times[]';

        timeInputContainer.appendChild(newTimeInput);
        timeFieldsContainer.appendChild(timeInputContainer);
    });

    form.addEventListener('submit', (event) => {
        event.preventDefault();

        const file = fileInput.files ? fileInput.files[0] : null;
        if (!file) {
            alert('Please upload an Excel file.');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = e.target?.result;
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            const formData = new FormData(form);
            const newRow: { [key: string]: any } = {};
            const times: string[] = [];

            formData.forEach((value, key) => {
                if (key === 'times[]') {
                    if (value) {
                        times.push(value.toString());
                    }
                } else if (key !== 'excel-file') { // Exclude the file input from the data
                    newRow[key] = value.toString();
                }
            });

            newRow['times'] = times.join(', ');

            // Order the data to match a consistent column order
            const rowData = [
                newRow['field1'],
                newRow['field2'],
                newRow['field3'],
                newRow['field4'],
                newRow['field5'],
                newRow['select-field'],
                newRow['number-field'],
                newRow['times']
            ];

            XLSX.utils.sheet_add_aoa(worksheet, [rowData], { origin: -1 });

            // Generate and trigger download of the modified file
            XLSX.writeFile(workbook, file.name);

            alert('The Excel file has been updated and downloaded!');
            form.reset();
        };
        reader.readAsArrayBuffer(file);
    });
});
