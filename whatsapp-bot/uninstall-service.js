var Service = require('node-windows').Service;

var svc = new Service({
  name: 'WareCore WhatsApp Bot',
  script: 'C:\\Users\\Administrator\\Desktop\\асер\\WareCore Reports\\whatsapp-bot\\bot.js'
});

svc.on('uninstall', function () {
  console.log('Service uninstalled');
});

svc.uninstall();
