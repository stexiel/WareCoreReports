var Service = require('node-windows').Service;

var svc = new Service({
  name: 'WareCore Telegram Bot',
  script: 'C:\\Users\\Administrator\\Desktop\\асер\\WareCore Reports\\telegram-bot\\bot.js'
});

svc.on('uninstall', function () {
  console.log('Service uninstalled');
});

svc.uninstall();
