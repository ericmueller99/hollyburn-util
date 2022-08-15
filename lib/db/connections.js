const sql = require('seriate');

const dbConnections = {
    checkEnv: (variable) => {
        if (!process.env[variable]) {
            throw new Error(`process.env.${variable} is missing`);
        }
    },
    residentPortal: {
        live: {
            server: process.env.DB_RESIDENTPORTAL_LIVE_SERVER,
            user: process.env.DB_RESIDENTPORTAL_LIVE_USER,
            port: 1433,
            password: process.env.DB_RESIDENTPORTAL_LIVE_PASSWORD,
            database: process.env.DB_RESIDENTPORTAL_LIVE_NAME,
            options: {
                encrypt: true
            }
        },
        test: {
            server: "sql.hollyburn.com",
            user: process.env.DB_RESIDENTPORTAL_TEST_USER,
            port: 1433,
            password: process.env.DB_RESIDENTPORTAL_TEST_PASSWORD,
            database: process.env.DB_RESIDENTPORTAL_TEST_NAME,
            options: {
                encrypt: true
            }
        },
    },
    yardi: {
        live: {
            server: "100.66.15.61",
            user: process.env.DB_YARDI_LIVE_USER,
            password: process.env.DB_YARDI_LIVE_PASSWORD,
            database: process.env.DB_YARDI_LIVE_NAME,
            requestTimeout: 30000
        },
        test: {
            server: "100.66.15.91\\F015DB91T_2K16",
            user: process.env.DB_YARD_TEST_USER,
            password: process.env.DB_YARDI_TEST_PASSWORD,
            database: process.env.DB_YARDI_TEST_NAME,
            requestTimeout: 30000,
        }
    },
    api: {
        live: {
            server: process.env.DB_HOLLYBURNAPI_SERVER,
            user: process.env.DB_HOLLYBURNAPI_USER,
            port: 1433,
            password: process.env.DB_HOLLYBURNAPI_PASSWORD,
            database: process.env.DB_HOLLYBURNAPI_NAME,
            options: {
                encrypt: true
            }
        }
    },
    getResidentPortal() {return this.residentPortal.live},
    getResidentPortalTest() {return this.residentPortal.test},
    getHollyburnApi() {return this.api.live},
    getYardi() {return this.yardi.live},
    getYardiTest() { return this.yardi.test},
    runQuery: function (connection, query, params = {}) {
        return new Promise((resolve, reject) => {
            sql.getPlainContext(connection)
                .step('query', {
                    query: query,
                    params: params
                })
                .end(results => {
                    if (results.query.length > 0) {
                        resolve(results.query);
                    }
                    else {
                        resolve([]);
                    }
                })
                .error(error => {
                    reject(error);
                })
        })
    }
}

module.exports = dbConnections;