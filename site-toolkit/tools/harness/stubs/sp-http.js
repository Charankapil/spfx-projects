const v1 = { name: 'v1', overrideWith: () => ({ name: 'search-v3' }) };
module.exports = { ODataVersion: { v3: '3.0', v4: '4.0' }, SPHttpClient: { configurations: { v1 } } };
