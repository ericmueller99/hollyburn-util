const sql = require('seriate');
const mssql = require('mssql');

const propertiesEntity = {
    defaultConnection: {
        user: process.env.DB_RESIDENTPORTAL_LIVE_USER,
        password: process.env.DB_RESIDENTPORTAL_LIVE_PASSWORD,
        database: process.env.DB_RESIDENTPORTAL_LIVE_NAME,
        server: process.env.DB_RESIDENTPORTAL_LIVE_SERVER,
        options: {
            encrypt: true
        }
    },
    getPropertyById(connection, propertyHMY) {
        return new Promise((resolve,reject) => {
            sql.getPlainContext(connection)
                .step('property', {
                    query: "select * from properties where property_hmy=@propertyHMY",
                    params: {
                        propertyHMY: {
                            type: sql.Int,
                            val: propertyHMY
                        }
                    }
                })
                .end(results => {
                    if (results.property.length> 0) {
                        resolve(results.property[0]);
                    }
                    else {
                        resolve({});
                    }
                })
                .error(error => {
                    reject(error);
                })
        })
    },
    getAllPropertiesWithAmenities(connection, whereStatement = '', propertiesParams = {}) {
        return new Promise((resolve, reject) => {
            const query = "select * from properties " + whereStatement;
            sql.getPlainContext(connection)
                .step('properties', {
                    query: query,
                    params: {
                        ...propertiesParams
                    }
                })
                .step('amenities', {
                    query: "select property_hmy, amenities_subcategories as amenity from properties_amenities_new\n" +
                        "inner join properties_amenities_categories pac on properties_amenities_new.category_id = pac.id\n" +
                        "inner join properties_amenities_subcategories pas on properties_amenities_new.amenity = pas.id"
                })
                .end(results => {

                    let {properties, amenities} = results;
                    if (properties.length <= 0) {
                        resolve([]);
                    }

                    //match amenities to properties
                    for (let property of properties) {
                        const propAmenities = amenities.filter(e => e.property_hmy === property.property_hmy);
                        property.amenities = propAmenities.map(e => e.amenity)
                    }
                    resolve(properties);

                })
                .error(error => {
                    reject(error);
                })
        })
    },
    async getPropertyBookingScheduleByEmail(emailAddress, connection = this.defaultConnection) {
        const pool = await mssql.connect(connection);
        const results = await pool.request()
            .input('emailAddress', mssql.NVarChar, emailAddress)
            .query("select bs.*, p.meeting_timezone as meeting_timezone, booking_default_duration\n" +
                "from booking_schedule bs\n" +
                "inner join properties p on p.property_hmy = bs.property_hmy\n" +
                "where p.rm_email = @emailAddress")
        return results.recordsets[0] || [];
    },
    async getPropertyBookingScheduleByPropertyHmy(propertyHMY, connection = this.defaultConnection) {
        const pool = await mssql.connect(connection);
        const results = await pool.request()
            .input('propertyHMY', mssql.Int, propertyHMY)
            .query("select bs.*, p.meeting_timezone as meeting_timezone, booking_default_duration, booking_unit_modifier, rm_email\n" +
                "from booking_schedule bs\n" +
                "inner join properties p on p.property_hmy = bs.property_hmy\n" +
                "where p.property_hmy=@propertyHMY")
        return results.recordsets[0] || [];
    }
}

module.exports = propertiesEntity