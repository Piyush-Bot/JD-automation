const { buildMenuCard } = require('./menuCard');
const { buildJdCreatFormCard } = require('./jdFormCard');
const { buildJdEditCard } = require('./editFetchForm');
const { buildEditFormCard } = require('./editFormCard');
const { buildJdResultCard } = require('./resultCard');
const { buildPopupCard, buildTaskModuleResponse } = require('./popupCard');

module.exports = {
    buildMenuCard,
    buildJdCreatFormCard,
    buildJdEditCard,
    buildEditFormCard,
    buildJdResultCard,
    buildPopupCard,
    buildTaskModuleResponse
};
