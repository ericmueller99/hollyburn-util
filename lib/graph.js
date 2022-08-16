//@ts-check

const sql = require('seriate');
const fs = require("fs");
const path = require("path");
const request = require("request");
const axios = require("axios").default;
const {detect} = require('mmmagic-type');
const propertiesEntity = require('./db/properties');
const dbConnections = require('./db/connections');
const {checkFolderExists: localFolderCheck} = require('./common');

//gets metadata for a file.  Must have an absoluteUrl that can be converted to a base64 sharing string.
async function getMetaData(graphClass) {
    return new Promise((resolve, reject) => {

        let url = graphClass.absoluteFileUrl;
        // console.log('Getting file Meta from SharePoint.  Absolute Url is: ' + graphClass.absoluteFileUrl);
        let buff = Buffer.from(url);
        let base64Value = buff.toString('base64');
        let encodedUrl = "u!" + base64Value.replace(/\//g, '_').replace(/\+/g, '-');
        // console.log('https://graph.microsoft.com/v1.0/shares/' + encodedUrl + '/driveItem?$expand=listItem')
        let requestOptions = {
            method: 'GET',
            headers: {
                Authorization: 'Bearer ' + graphClass.bearerToken
            },
            uri: 'https://graph.microsoft.com/v1.0/shares/' + encodedUrl + '/driveItem',
            json: true
        }

        request(requestOptions, (error, response, body) => {
            if (error) {
                reject(error);
            }
            else {
                if (body.error) {
                    resolve(body);
                }
                else {
                    resolve(body);
                }
            }
        })
    })
}

//gets metadata with $expand=listItem on.  Only works for certain files/folders.
async function getMetaDataListItems(graphClass) {
    return new Promise((resolve, reject) => {
        let url = graphClass.absoluteFileUrl;
        // console.log('Getting file Meta from SharePoint.  Absolute Url is: ' + graphClass.absoluteFileUrl);
        let buff = Buffer.from(url);
        let base64Value = buff.toString('base64');
        let encodedUrl = "u!" + base64Value.replace(/\//g, '_').replace(/\+/g, '-');
        // console.log('https://graph.microsoft.com/v1.0/shares/' + encodedUrl + '/driveItem?$expand=listItem')
        let requestOptions = {
            method: 'GET',
            headers: {
                Authorization: 'Bearer ' + graphClass.bearerToken
            },
            uri: 'https://graph.microsoft.com/v1.0/shares/' + encodedUrl + '/driveItem?$expand=listItem',
            json: true
        }

        request(requestOptions, (error, response, body) => {
            if (error) {
                reject(error);
            }
            else {
                if (body.error) {
                    resolve(body);
                }
                else {
                    resolve(body);
                }
            }
        })
    })

}

//downloads the file from SharePoint and saves it in the temp folder.
async function downloadFile(graphClass) {
    return new Promise((resolve,reject) => {
        let requestOptions = {
            uri: graphClass.downloadUrl,
            contentType: graphClass.fileMimeType,
            method: 'GET'
        }
        request(requestOptions)
            .on('response', res => {
                delete res.headers;
            })
            .on('end', () => {
                resolve();
            })
            .on('error', error => {
                reject(error);
            })
            .pipe(fs.createWriteStream(path.join(__dirname, '../temp/' + graphClass.fileName)))
    })
}

//authenticates or gets an existing authentication to SharePoint
async function authenticate(graphClass) {
    return new Promise((resolve,reject) => {

        //get a new token
        let getNewToken = function () {

            let requestOptions = {
                uri: "https://login.microsoftonline.com/eb40f7e6-4744-450e-aee0-d66ca1450f96/oauth2/token",
                method: "POST",
                formData: {
                    client_id: "c0353ac7-8cdd-41b0-bce0-79101898e11f",
                    resource: "https://graph.microsoft.com",
                    client_secret: "HK7/PxhzvZ4hynutfL6kavN0DM0tVfaurtewkVHnpXs=",
                    grant_type: "client_credentials"
                },
                json: true
            }
            request(requestOptions, function (error, response,body) {
                if (error) {
                    reject(error);
                }
                else {

                    //saving the token in the database
                    sql.getPlainContext(dbConnections.getResidentPortal())
                        .step('saveToken', {
                            query: "update auth_tokens set token=@token, expires=DATEADD(minute, 55, getdate()) where api='microsoft-graph' " +
                                "if @@ROWCOUNT=0 " +
                                "insert into auth_tokens (api, token, expires) values ('microsoft-graph', @token, DATEADD(minute, 55, getdate()))",
                            params: {
                                token: {
                                    type: sql.varChar,
                                    val: body.access_token
                                }
                            }
                        })
                        .error(function (error) {
                            reject(error);
                        });

                    graphClass.bearerToken = body.access_token;
                    resolve();

                }
            })

        }

        //get an existing token
        let getExistingToken = function () {
            sql.getPlainContext(dbConnections.getResidentPortal())
                .step('token', {
                    query: "select token from auth_tokens where api='microsoft-graph' and expires > GETDATE()"
                })
                .end(function (results) {
                    if ((results.token[0]) && (results.token[0].token)) {
                        graphClass.bearerToken = results.token[0].token;
                        resolve();
                    }
                    //no valid token.  getting a new one
                    else {
                        getNewToken();
                    }
                })
                .error(function (error) {
                    reject(error);
                })
        }

        //starting the callback chain to get a token.  When one is retrieved then the function is resolved and this.bearerToken is set
        getExistingToken();

    })
}

//gets a list of files inside of a folder.
async function getFolderFiles(graphClass) {
    return new Promise((resolve,reject) => {
        let url = graphClass.absoluteFileUrl;
        // console.log('Getting files inside of folder: ' + url);
        let buff = Buffer.from(url);
        let base64Value = buff.toString('base64');
        let encodedUrl = "u!" + base64Value.replace(/\//g, '_').replace(/\+/g, '-');
        // console.log('The base64 value for this folder is: ' + base64Value);
        let requestOptions = {
            method: 'GET',
            headers: {
                Authorization: 'Bearer ' + graphClass.bearerToken
            },
            uri: 'https://graph.microsoft.com/v1.0/shares/' + encodedUrl + '/driveItem/children',
            json: true
        }
        request(requestOptions, (error, response, body) => {
            if (error) {
                reject(error);
            }
            else {
                if (body.error) {
                    reject(body);
                }
                else {
                    resolve(body);
                }
            }
        })
    })
}

//moves a file from a folder to a new folder.
async function moveAllFilesFromFolderToFolder(graphClass, filesToMove) {
    return new Promise((resolve,reject) => {
        try {

            //checking if all moves have been completed
            let moveCallback = function (body, error, fileId) {

                console.log('this is move callback');
                if (error) {
                    console.log('There was an error: ' + error);
                }
                else {
                    console.log(body);
                }

                let allFilesCompleted = true;
                for (let i in filesToMove) {
                    if (filesToMove[i].fileId === fileId) {
                        filesToMove[i].completed = true;

                        //appending the result
                        if (error) {
                            filesToMove[i].results = {
                                result: false,
                                error: error
                            }
                        }
                        else {
                            filesToMove[i].results = {
                                result: true,
                                details: body
                            }
                        }

                    }
                    //is a file that does not match the current Id not completed?
                    if (!filesToMove[i].completed) {
                        allFilesCompleted = false;
                    }

                    //End of Loop
                    if (parseInt((i)+1) === (parseInt(filesToMove.length))) {
                        if (allFilesCompleted) {
                            resolve(filesToMove);
                        }
                    }
                }
            }

            //looping through the filesToMove array.
            for (let i in filesToMove) {
                let requestOptions = {
                    method: 'PATCH',
                    uri: 'https://graph.microsoft.com/v1.0/drives/' + filesToMove[i].fileDriveId + '/items/' + filesToMove[i].fileId,
                    headers: {
                        Authorization: 'Bearer ' + graphClass.bearerToken
                    },
                    body: {
                        parentReference: {
                            driveId: filesToMove[i].targetFolderDriveId,
                            id: filesToMove[i].targetFolderId
                        }
                    },
                    json: true
                }
                request(requestOptions, (error, response, body) => {
                    moveCallback(body, error, filesToMove[i].fileId);
                })
            }

        }
        catch (error) {
            reject(error);
        }
    })
}

class Graph {

    #endpoints = {
        metadata: 'https://graph.microsoft.com/v1.0/shares/{{encodedUrl}}/driveItem',
        createFolder: "https://graph.microsoft.com/v1.0/drives/{{driveId}}/items/{{parentFolderId}}/children",
        uploadFile: "https://graph.microsoft.com/v1.0/drives/{{driveId}}/items/{{folderId}}:/{{documentName}}:/content",
        calendarSchedule: "https://graph.microsoft.com/v1.0/users/{{emailAddress}}/calendar/getSchedule",
        sendCalendarInvite: "https://graph.microsoft.com/v1.0/users/{{fromEmail}}/calendar/events",
        getCalendarInvite: "https://graph.microsoft.com/v1.0/users/rent@hollyburn.com/calendar/events/{{eventId}}/",
        cancelCalendarInvite: "https://graph.microsoft.com/v1.0/users/rent@hollyburn.com/calendar/events/{{eventId}}/cancel"
    }
    #residentPortalDbConnection;
    #yardiDbConnection;
    bearerToken;
    absoluteFileUrl;
    downloadUrl;
    fileMimeType;
    fileName;

    constructor(options = {}) {

        //making sure that the temp folder exists.
        localFolderCheck(path.join(__dirname, "../temp"));

        console.log(dbConnections);

        //is this a test?
        const {isTest = null} = options
        this.#residentPortalDbConnection = isTest ? dbConnections.getResidentPortalTest() : dbConnections.getResidentPortal();
        this.#yardiDbConnection = isTest ? dbConnections.getYardiTest() : dbConnections.getYardi();

    }

    //sends a request using Axios
    #sendRequest(url, data = {}, methodType = 'post', headers = {}, requestContext) {
        return new Promise((resolve,reject) => {

            const methodTypes = new Set(['post', 'get', 'put', 'delete'])
            if (!methodTypes.has(methodType.toLowerCase())) {
                reject(new Error('methodType is not valid'));
            }

            axios({
                method: methodType,
                url: url,
                data,
                headers: {
                    Authorization: 'Bearer ' + this.bearerToken,
                    ...headers
                }
            })
                .then(res => {
                    if (res.status === 200) {
                        res.data.result = true;
                        resolve(res.data);
                    }
                    else if (res.status === 201) {
                        res.data.result = true;
                        resolve(res.data);
                    }
                    else if (res.status === 202) {
                        res.data.result = true;
                        resolve(res.data);
                    }
                    else {
                        console.log('res was not 200 or 201');
                        console.log(res);
                        reject(new Error('res was not 200 or 201'));
                    }
                })
                .catch(error => {
                    reject(error);
                })
        })
    }

    //checks if a folder path exists in SharePoint.
    async checkIfFolderExists(folderPath) {
        try {
            if (folderPath) {

                this.absoluteFileUrl = folderPath;

                //authenticate
                await authenticate(this);

                //getting meta data for the folder.  If no meta data is found then the folder does not exist
                let metaData = await getMetaData(this);
                if (metaData.name) {
                    metaData.result = true;
                    return metaData;
                }
                else {
                    metaData.result = false;
                    return metaData;
                }

            }
            else {
                throw new Error('folderPath is a required field.');
            }
        }
        catch (error) {
            throw error;
        }
    }

    //gets file metadata with expanded ListItems
    async getFileMetaData(filePath) {
        try {
            if (filePath) {
                this.absoluteFileUrl = filePath;
                await authenticate(this);
                let metaData = await getMetaDataListItems(this);
                metaData.result = !!metaData.name;
                return metaData;
            }
            else {
                throw new Error('filePath is required');
            }
        }
        catch (error) {
            throw error;
        }
    }

    //gets a document
    async getDocument(documentPath, documentName, callback = null) {
        try {

            if (!documentPath || !documentName) {
                return callback ? callback(new Error('documentPath and documentName are required')) : Promise.reject(new Error("documentPath and documentName are required"));
            }

            this.absoluteFileUrl = documentPath + '/' + documentName;
            this.fileName = documentName;

            //getting the bearer token
            await authenticate(this)

            const url = this.#createUrl(this.#endpoints.metadata, `${documentPath}/${documentName}`);
            const metaData = await this.#sendRequest(url, null, 'get');

            //if metadata returned the data we need to download the file then download it.
            if ((metaData['@microsoft.graph.downloadUrl']) && (metaData.file)) {

                //download file name
                this.downloadUrl = metaData['@microsoft.graph.downloadUrl'];
                this.fileMimeType = metaData.file.mimeType;

                await downloadFile(this);

                let {mimeType} = metaData?.file?.mimeType;
                if (!mimeType) {
                    const r = await detect(fs.readFileSync(path.join(__dirname, `../temp/${documentName}`)))
                    mimeType = r.mime;
                }

                return {
                    contentType: mimeType
                }

            }
            //meta data call did not return require info to download file.  throw error.
            else {
                return callback ? callback(new Error('Could not get downloadUrl and/or file mimeType.')) : Promise.reject(new Error('Could not get downloadUrl and/or file mimeType.'));
            }

        }
        catch (error) {
            return callback ? callback(error) : Promise.reject(error);
        }
    }

    //gets a document with the download Url generated from the Microsoft Graph MetaData endpoint.
    async getDocumentWithDownloadUrl(downloadUrl, mimeType) {
        try {
            if ((downloadUrl) && (mimeType)) {

                //authenticating to Graph.
                await authenticate(this);

                //downloading the file using downloadFile function
                this.downloadUrl = downloadUrl;
                this.fileMimeType = mimeType;
                await downloadFile(this)
                    .then(results => {
                        return results;
                    })
                    .catch(error => {
                        if ((error.message) && (error.stack)) {
                            throw error;
                        }
                        else {
                            throw new Error('Graph had an error downloading the file.  The error was: ' + JSON.stringify(error));
                        }
                    })

            }
            else {
                throw new Error('downloadUrl and mimeType are required.');
            }
        }
        catch (error) {
            throw error;
        }
    }

    //copies all files in a parent folder to the target driveId & folderId
    async moveFolderFilesToFolder(copyFromFolderPath, copyToFolderPath) {
        try {
            if ((copyFromFolderPath) && (copyToFolderPath)) {

                await authenticate(this);

                //getting the files inside of the copyFromFolderPath
                this.absoluteFileUrl = copyFromFolderPath;
                let folderFiles = await getFolderFiles(this);
                if (!folderFiles.value) {
                    throw new Error('Unable to find copyFromFolderPath on the SharePoint site.');
                }

                //getting the target folder Id from the metadata lookup function.
                this.absoluteFileUrl = copyToFolderPath;
                let targetFolder = await getMetaData(this)
                if (!targetFolder.name) {
                    throw new Error('Unable to find meta data for target folder.');
                }

                //looping through the files in the copyFromFolderPath and moving them to the target folder
                let filesToMove = [];
                for (let i in folderFiles.value) {
                    filesToMove.push({
                        fileDriveId: folderFiles.value[i].parentReference.driveId,
                        fileId: folderFiles.value[i].id,
                        targetFolderDriveId: targetFolder.parentReference.driveId,
                        targetFolderId: targetFolder.id
                    })
                    if (parseInt((i)+1) === (parseInt(folderFiles.value.length))) {
                        console.log('End of Loop');
                        return await moveAllFilesFromFolderToFolder(this, filesToMove);
                    }
                }

            }
            else {
                throw new Error('copyFromFolderPath and copyToFolderPath are both required fields.');
            }
        }
        catch (error) {
            throw error;
        }
    }

    //gets a list of documents from a folder
    async getDocumentsInFolder(documentPath) {
        try {
            if (documentPath) {

                //authenticating to SharePoint
                await authenticate(this);

                this.absoluteFileUrl = documentPath;
                return await getFolderFiles(this);

            }
            else {
                throw new Error('documentPath is required.');
            }
        }
        catch (error) {
            if (!error.stack) {
                throw new Error(JSON.stringify(error));
            }
            else {
                throw error;
            }
        }
    }

    //uploads a document
    async uploadDocument(targetPath, documentName, tempFile, callback = null) {
        try {

            if (!targetPath || !documentName) {
                return callback ? callback(new Error('targetPath and documentName are required fields')) : Promise.reject(new Error('targetPath and documentName are required fields'));
            }

            //auth
            await authenticate(this);

            //getting metadata for the target path.
            const targetMetadata = await this.#sendRequest(this.#createUrl(this.#endpoints.metadata, targetPath), {}, 'get');
            const {driveId} = targetMetadata?.parentReference;
            const {id: folderId} = targetMetadata;

            //driveId and id returned?
            if (!driveId || !folderId) {
                console.log('driveId or folderId is missing');
                return callback ? callback(new Error('Unable to get driveId and id from target metadata call.  Check this is a valid URL')) : Promise.reject(new Error('Unable to get driveId and id from target metadata call.  Check this is a valid URL'));
            }

            const encodedDocumentName = encodeURI(documentName);
            const uploadUri = this.#createUrl(this.#endpoints.uploadFile, {
                driveId,
                folderId,
                documentName: encodedDocumentName
            });

            console.log('got here graph');
            console.log(uploadUri);

            return await this.#sendRequest(uploadUri, fs.readFileSync(path.join(__dirname, '../temp/' + tempFile)), 'put', {
                "Content-Type": "application/pdf"
            });

        }
        catch(error) {
            return callback ? callback(error) : Promise.reject(error);
        }

    }

    #createUrl(endpoint, variables, encodeString = true) {

        //TODO write something in here that finds all the variables in the endpoint string and warns the user if one is not replaced.

        let uri = endpoint;

        const encodeUrl = (value) => {
            let buff = Buffer.from(encodeURI(value));
            let base64Value = buff.toString('base64');
            return "u!" + base64Value.replace(/\//g, '_').replace(/\+/g, '-');
        }

        if (typeof variables === 'object') {
            Object.keys(variables).forEach(key => {
                if (variables[key].encode) {
                    variables[key] = encodeUrl(variables[key])
                }
                uri = uri.replace('{{' + key + '}}', variables[key]);
            })
            return uri;
        }
        else if (typeof variables === 'string') {
            let encodedUrl = variables;
            if (encodeString) {
                encodedUrl = encodeUrl(encodedUrl);
            }
            return uri.replace('{{encodedUrl}}', encodedUrl);
        }


    }

    //creates a folder in SharePoint.  Requires a parent folder/path for this to work.
    async createFolder(parentPath, folderName, callback = null) {
        try {

            //if either parentPath or folderName is nothing then return an error.
            if (!parentPath || !folderName) {
                return callback ? callback(new Error('parentPath and folderName are required fields')) : Promise.reject(new Error('parentPath and folderName are required fields'));
            }

            //making sure that the folder does not have extra spaces in it.
            folderName = folderName.trimStart().trimEnd();

            //authenticating
            await authenticate(this);

            //getting data for parentPath
            const parentMetaData = await this.#sendRequest(this.#createUrl(this.#endpoints.metadata, parentPath), {}, 'get');
            const {driveId} = parentMetaData.parentReference;
            const {id: parentFolderId} = parentMetaData;

            if (!parentFolderId || !driveId) {
                return callback ? callback(new Error('unable to get parentPath metadata.  Check that its a valid Url')) : Promise.reject(new Error('unable to get parentPath metadata.  Check that its a valid Url'));
            }

            //generating Url
            const url = this.#endpoints.createFolder.replace('{{driveId}}', driveId).replace('{{parentFolderId}}', parentFolderId)
            const data = {
                name: folderName,
                folder: {},
                "@microsoft.graph.conflictBehavior" : "fail" //make sure that if the folder already exists do nothing
            }

            //sending the folder create request
            try {
                let res = await this.#sendRequest(url, data, 'post');
                res.result = true;
                res.message = "folder created";
                return res;
            }
            catch (error) {
                const {status, statusText} = error.response;
                if (status === 409 && statusText === 'Conflict') {
                    return {result:true, message: "folder already exists"};
                }
                return Promise.reject(error);
            }

        }
        catch (error) {
            return callback ? callback(error) : Promise.reject(error);
        }
    }

    //sets the metadata for an item.
    async setMetaData(tagData, target, targetMeta = null) {
        try {

            //auth
            await authenticate(this);

            let metaData;
            if (!targetMeta) {
                this.absoluteFileUrl = target;
                metaData = await getMetaData(this);
            }
            else {
                metaData = targetMeta;
            }

            let requestOptions = {
                method: 'PATCH',
                uri: "https://graph.microsoft.com/v1.0/drives/" + metaData.parentReference.driveId + "/items/" + metaData.id + "/listItem/fields",
                headers: {
                    Authorization: 'Bearer ' + this.bearerToken
                },
                body: tagData,
                json: true
            }
            request(requestOptions, (error, response, body) => {
                if (error) {
                    console.log(error);
                    return Promise.reject(error);
                    throw error;
                }
                else {
                    console.log(body);
                    return Promise.resolve();
                }
            })
        }
        catch (error) {
            console.log(error);
            return Promise.reject();
        }
    }

    //gets and returns calendar availability for a building from their email address
    async getCalendarAvailability(propertyHMY, startDate, endDate, numberOfSuites, options = {useDifferenceAsInterval: false}) {

        await authenticate(this);

        const {useDifferenceAsInterval} = options;

        //getting all days between start and end date
        let startDateSlice = new Date(startDate);
        const bookingDays = [];
        while (startDateSlice.getTime() <= endDate.getTime()) {
            bookingDays.push(startDateSlice);
            startDateSlice = new Date(startDateSlice.getTime() + 86400000);
        }

        //getting property / schedule information
        const schedule = await propertiesEntity.getPropertyBookingScheduleByPropertyHmy(propertyHMY, this.#residentPortalDbConnection);

        if (schedule.length < 1) {
            return Promise.reject(new Error('there is no schedule setup for this building.'));
        }

        //these variables are at the property level, but instead of querying twice we are just including them in the query.
        const {booking_default_duration, meeting_timezone, booking_unit_modifier, rm_email} = schedule[0];

        //splitting the start date and end date into booking windows which will then be filtered out based on the schedule and the calendar availability.
        let timeIntervals = [];
        let timeIntervalStart = new Date(startDate);
        const timeIntervalInMillis =  useDifferenceAsInterval ? endDate.getTime() - startDate.getTime() : numberOfSuites === 1 ? booking_default_duration * 60000 : (((booking_default_duration * booking_unit_modifier - booking_default_duration) * numberOfSuites) + booking_default_duration) * 60000
        let timeIntervalEnd = new Date(startDate.getTime() + timeIntervalInMillis);
        while (timeIntervalStart <= endDate) {
            timeIntervals.push({
                start: timeIntervalStart,
                end: timeIntervalEnd,
                availabilityInt: null
            })
            timeIntervalStart = new Date(timeIntervalEnd);
            timeIntervalEnd = new Date(timeIntervalStart.getTime() + timeIntervalInMillis);
        }

        //getting the availability from Graph for this calendar.
        const url = this.#createUrl(this.#endpoints.calendarSchedule, {emailAddress: rm_email}, false);
        const postData = {
            schedules: [rm_email],
            startTime: {
                dateTime: startDate,
                timeZone: meeting_timezone
            },
            endTime: {
                dateTime: endDate,
                timeZone: meeting_timezone
            },
            availabilityViewInterval: timeIntervalInMillis / 60000
        }
        const calendar = await this.#sendRequest(url, postData);
        const [{availabilityView}] = calendar.value

        //assigning the availability view to the time intervals array.
        for (const i in timeIntervals) {
            timeIntervals[i].availabilityInt = availabilityView[i]
        }

        //removing time intervals that don't align with the property schedule or the graph calendar schedule.
        const weekdays = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        let timeIntervalsFilteredBySchedule = [];
        for (const day of bookingDays) {
            const dayOfWeek = weekdays[day.getDay()].toLowerCase();
            for (const s of schedule) {
                const startDate = new Date(day.setHours(s[`${dayOfWeek}_start`].slice(0,2)));
                const endDate = new Date(day.setHours(s[`${dayOfWeek}_end`].slice(0,2)));
                const filteredBySchedule = timeIntervals.filter(t => t.start >= startDate && t.end <= endDate && t.availabilityInt === "0");
                timeIntervalsFilteredBySchedule = [...timeIntervalsFilteredBySchedule, ...filteredBySchedule];
            }
        }

        return timeIntervalsFilteredBySchedule;

    }

    //sends a calendar invite
    async sendCalendarInvite(subject, startDate, endDate, fromEmail, toEmails, templateHtml, timeZone, location) {

        if (!Array.isArray(toEmails)) {
            throw new Error('toEmails must be an array.');
        }

        //authenticate to graph
        await authenticate(this);

        //creating the URL
        const url = this.#createUrl(this.#endpoints.sendCalendarInvite, {fromEmail}, false);

        //creating the attendees array and the body of the request
        const attendees = [];
        for (const email of toEmails) {
            attendees.push({
                emailAddress: {
                    address: email,
                },
                type: "required"
            })
        }
        const postData = {
            subject: subject,
            body: {
                contentType: "HTML",
                content: templateHtml
            },
            start: {
                dateTime: startDate,
                timeZone: timeZone
            },
            end: {
                dateTime: endDate,
                timeZone: timeZone
            },
            location: {
                displayName: location
            },
            attendees: attendees
        }
        return await this.#sendRequest(url, postData);

    }

    //gets a calendar event from ID
    async getCalendarEventDetails(eventId) {

        await authenticate(this);
        const url = this.#createUrl(this.#endpoints.getCalendarInvite, {eventId}, false);
        const results = await this.#sendRequest(url, null, 'get');
        return results || {}

    }

    //cancel a calendar event from ID
    async cancelCalendarEvent(eventId) {

        await authenticate(this);
        const url = this.#createUrl(this.#endpoints.cancelCalendarInvite, {eventId}, false);
        const results = await this.#sendRequest(url, null, 'post');
        console.log(results);

    }


}

module.exports = {
    Graph
}
