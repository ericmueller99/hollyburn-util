const fs = require('fs');

const checkFolderExists = (folder) => {
    if (!fs.existsSync(folder)) {
        fs.mkdirSync(folder);
    }
}

module.exports = {
    checkFolderExists
}