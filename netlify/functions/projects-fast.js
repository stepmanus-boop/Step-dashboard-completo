const projects = require('./projects');

exports.handler = async (event = {}, context = {}) => {
  const queryStringParameters = {
    ...(event.queryStringParameters || {}),
    preferCache: '1',
  };
  delete queryStringParameters.force;

  return projects.handler({
    ...event,
    queryStringParameters,
  }, context);
};
