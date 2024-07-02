const XLSX = require('xlsx');

const getConfig = () => {
    return {
        user: "admin",
        pass: "Q%21BmT3bpHID5pr",
        urlBase: "https://bpms.abbott.com:14503",
        clientId: "XCLNYNJPICJNJRTVTGAHMEDCSYHVPCHW",
        clientSecret: "398049202667f15b20aa403056718293",
        workspace: "workflow"
    };
}

const getToken = async () => {
    // Obtener la configuración
    const { user, pass, urlBase, clientId, clientSecret, workspace } = getConfig();
    // Obtener token
    const endpoint = `${urlBase}/api/1.0/${workspace}/oauth2/token`;
    const myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/json");

    const raw = JSON.stringify({
        "grant_type": "password",
        "scope": "*",
        "client_id": clientId,
        "client_secret": clientSecret,
        "username": user,
        "password": pass
    });

    const requestOptions = {
        method: "POST",
        headers: myHeaders,
        body: raw
    };

    try {
        const response = await fetch(endpoint, requestOptions);

        if (!response.ok) {
            throw new Error(`Error: ${response.status} ${response.statusText}`);
        }

        const result = await response.json();
        return result.access_token;
    } catch (error) {
        console.error('Fetch error:', error);
        throw error; // Re-throw the error if needed
    }
}
const createUser = async (accessToken, user) => {
    
    try {
        const { urlBase, workspace } = getConfig();
        const endpoint = `${urlBase}/api/1.0/${workspace}/user`;
        
        const raw = JSON.stringify({
            "usr_username": user.username,
            "usr_firstname": user.firstName,
            "usr_lastname": user.lastName,
            "usr_email": user.email,
            "usr_due_date": "2030-12-31",
            "usr_status": "ACTIVE",
            "usr_role": user.role == 'Operator' ? "PROCESSMAKER_OPERATOR" : '',
            "usr_new_pass": user.password,
            "usr_cnf_pass": user.password,
            "usr_address": "",
            "usr_zip_code": "46135",
            "usr_country": "US",
            "usr_city": "",
            "usr_location": "",
            "usr_phone": "",
            "usr_position": "",
            "usr_calendar": "00000000000000000000000000000001"
        });
        
        const requestOptions = {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${accessToken}`
            },
            body: raw
        };

        const response = await fetch(endpoint, requestOptions);

        if (!response.ok) {
            throw new Error(`Error: ${response.status} ${response.statusText}`);
        }

        const result = await response.json();
        console.log('Usuario creado:', result); // Maneja la respuesta como necesites
    } catch (error) {
        console.error('Error al crear usuario:', error);
    }
}

const readXlsx = (path) => {
    // Carga el archivo Excel
    const workbook = XLSX.readFile(path);

    // Selecciona la primera hoja
    //const sheetName = workbook.SheetNames['usuarios'];
    const worksheet = workbook.Sheets['usuarios'];

    // Convierte la hoja a JSON
    const data = XLSX.utils.sheet_to_json(worksheet);

    const processedUsers = data.map(user => {
        return {
            email: user['Correo electrónico'],
            firstName: user['Primer nombre'].trim(),
            lastName: user['Primer apellido'],
            username: user['Nombre del usuario'],
            password: user['Contraseña'],
            role: user['Rol ProcessMaker'],
            position: user.Cargo,
            mainSociety: user.Sociedad,
            society1: user['Sociedad 1'],
            society2: user['Sociedad 2']
        };
    });

    return processedUsers;

}

const main = async () => {
    try {

        const dataXls = readXlsx('C:/Users/juanp/OneDrive/Documentos/mi-proyecto/Matriz de Usuarios - Clientes.xlsx');
        const accessToken = await getToken();
        for (let user of dataXls) {
            await createUser(accessToken,user);
        }
        console.log(accessToken)

    } catch (error) {
        console.error('Error obteniendo el token:', error);
    }
}

main();

// En este caso, también asegúrate de manejar la lógica dentro de main() debido a la naturaleza asíncrona de await


return;













console.log(processedUsers);
