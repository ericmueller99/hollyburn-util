const {Graph} = require('./graph');

const getCalendarInvite = () => {

    const graph = new Graph();
    graph.getCalendarEventDetails('AAMkADk2NWVjZjk2LTgxZjktNDhiYy1hOTkyLTZlYmFhMGMyZjkzOQBGAAAAAAAQIqOnHg7xQaaiMA-frPe0BwBTCl3d1VeJR76Bvqd5JGweAAAAAAENAABTCl3d1VeJR76Bvqd5JGweAAFUEYr_AAA=')
        .then(res => {
            console.log(res);
        })
        .catch(error => {
            console.log(error);
        })
}

getCalendarInvite();