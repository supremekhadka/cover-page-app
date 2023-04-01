
/* =============== event Handler ============= */
form.addEventListener("submit", (e)=>{
    e.preventDefault();
    generate(currentData);
});
/* =========================================== */

const templateFileMap = {
    DL : "./templates/DL.docx",
    "C Programming": "./templates/CProgramming.docx",
    IIT: "./templates/DL.docx",
}

function loadFile(url, callback) {
    PizZipUtils.getBinaryContent(url, callback);
}

function generate(obj) {
    loadFile(
        templateFileMap[obj.subject],
        function (error, content) {
            if (error) {
                throw error;
            }
            var zip = new PizZip(content);
            var doc = new window.docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
            });

            // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
            doc.render(obj);

            var blob = doc.getZip().generate({
                type: "blob",
                mimeType:
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                // compression: DEFLATE adds a compression step.
                // For a 50MB output document, expect 500ms additional CPU time
                compression: "DEFLATE",
            });
            // Output the document using Data-URI
            saveAs(blob, "output.docx");
        }
    );
};