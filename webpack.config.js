const path = require('path')
const webpackConfig = require('@nextcloud/webpack-vue-config')

webpackConfig.entry = {
	admin: path.join(__dirname, 'src', 'admin.js'),
	personal: path.join(__dirname, 'src', 'personal.js'),
}

webpackConfig.output.filename = 'nc_ms365_calendar-[name].js'

module.exports = webpackConfig
