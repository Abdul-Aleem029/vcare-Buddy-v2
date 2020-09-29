var sql = require('mssql');

var dbConfig = {
    server: "vcare.database.windows.net",
    database: "vCareBud",
    user: "vcare",
    password: "Infosys@123",
    port: 1433
}

const getPersonas = async (name) => {
    try {
        console.log(name);
        // make sure that any items are correctly URL encoded in the connection string
        await sql.connect('mssql://vcare:Infosys@123@vcare.database.windows.net/vCareBud?encrypt=true')
        const result = await sql.query`select * from vcare.persona where name like ${name}`
        console.dir(result.recordset[0])
        return result.recordset[0];
    } catch (err) {
        // ... error checks
    }
}

export default getPersonas;