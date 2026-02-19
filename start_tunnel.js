const fs = require('fs');
const path = require('path');
const localtunnel = require('localtunnel');

(async () => {
  const tunnel = await localtunnel({ port: 8091 });
  const out = path.join(process.cwd(), 'tunnel_url.txt');
  fs.writeFileSync(out, tunnel.url, 'utf8');
  console.log(`TUNNEL_URL=${tunnel.url}`);
  tunnel.on('close', () => process.exit(0));
})();
