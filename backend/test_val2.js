const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');

async function run() {
    try {
        const csv = `Call Date,,Disposition,,Total Time,Talk Time,Ring Time,Caller ID,Target Name,Target Number
7/28/2026 13:07,,ANSWER,,'00:06:53,'00:06:46,'00:00:07,13175578549,Luk 1707,18332251707
7/28/2026 13:05,,MISSED,,'00:00:08,'00:00:00,'00:00:08,18047201389,Luk 1707,18332251707
7/28/2026 13:03,,ANSWER,,'00:01:28,'00:01:07,'00:00:21,16023775134,Luk 1707,18332251707`;
        fs.writeFileSync('test.csv', csv);
        
        const form = new FormData();
        form.append('sheet', fs.createReadStream('test.csv'));

        const response = await axios.post('http://localhost:3001/api/validate-numbers', form, {
            headers: form.getHeaders(),
        });
        console.log("Success:", response.data.stats);
    } catch (e) {
        console.error("Failed:", e.response ? e.response.data : e.message);
    }
}
run();
