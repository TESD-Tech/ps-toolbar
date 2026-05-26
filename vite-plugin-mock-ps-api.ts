/**
 * Vite plugin: mock PowerSchool REST API for local development.
 * Intercepts /ws/schema/table/U_TESD_PS_TOOLBAR_ICONS endpoints
 * and returns in-memory mock data so the admin page works in dev mode.
 */
export function mockPsApi() {
  // In-memory data store — persists across requests during dev session
  const store: Record<number, any> = {
    1001: {
      dcid: 1001,
      id: 'messages',
      icon: 'mail',
      href: '/admin/messages.html',
      title: 'Messages',
      description: 'View your school messages',
      count_sql: "SELECT COUNT(*) FROM U_TESD_MESSAGES WHERE recipient_id = ~(curuserid) AND read_flag = '0'",
      sort_order: 10,
      disabled: '0',
    },
    1002: {
      dcid: 1002,
      id: 'alerts',
      icon: 'bell',
      href: '/admin/alerts.html',
      title: 'Alerts',
      description: 'View system notifications and alerts',
      count_sql: "SELECT COUNT(*) FROM U_TESD_ALERTS WHERE status = 'active' AND (target_user = ~(curuserid) OR target_user IS NULL)",
      sort_order: 20,
      disabled: '0',
    },
    1003: {
      dcid: 1003,
      id: 'approvals',
      icon: 'star',
      href: '/admin/approvals.html',
      title: 'Approvals',
      description: 'Pending approvals requiring your review',
      count_sql: "SELECT COUNT(*) FROM U_TESD_APPROVALS WHERE approver_id = ~(curuserid) AND status = 'pending'",
      sort_order: 30,
      disabled: '0',
    },
  };
  let nextDcid = 1004;

  return {
    name: 'mock-ps-api',
    configureServer(server: any) {
      server.middlewares.use('/ws/schema/table/U_TESD_PS_TOOLBAR_ICONS', (req: any, res: any, next: any) => {
        // CORS headers
        res.setHeader('Content-Type', 'application/json');

        // Parse the URL to determine the record DCID
        const url = new URL(req.url, `http://${req.headers.host}`);
        const pathParts = url.pathname.split('/').filter(Boolean);
        const recordDcid = pathParts.length > 4 ? parseInt(pathParts[pathParts.length - 1], 10) : null;

        switch (req.method) {
          case 'GET': {
            if (recordDcid) {
              // GET by DCID
              const record = store[recordDcid];
              if (record) {
                res.statusCode = 200;
                res.end(JSON.stringify(record));
              } else {
                res.statusCode = 404;
                res.end(JSON.stringify({ error: 'Record not found' }));
              }
            } else {
              // GET all records
              const records = Object.values(store).filter((r: any) => r);
              res.statusCode = 200;
              res.end(JSON.stringify(records));
            }
            break;
          }

          case 'POST': {
            let body = '';
            req.on('data', (chunk: string) => { body += chunk; });
            req.on('end', () => {
              try {
                const data = JSON.parse(body);
                const dcid = nextDcid++;
                store[dcid] = { ...data, dcid };
                res.statusCode = 201;
                res.end(JSON.stringify(store[dcid]));
              } catch (e: any) {
                res.statusCode = 400;
                res.end(JSON.stringify({ error: e.message }));
              }
            });
            break;
          }

          case 'PUT': {
            if (!recordDcid) {
              res.statusCode = 400;
              res.end(JSON.stringify({ error: 'DCID required for update' }));
              break;
            }
            let body = '';
            req.on('data', (chunk: string) => { body += chunk; });
            req.on('end', () => {
              try {
                const data = JSON.parse(body);
                if (store[recordDcid]) {
                  store[recordDcid] = { ...store[recordDcid], ...data, dcid: recordDcid };
                  res.statusCode = 200;
                  res.end(JSON.stringify(store[recordDcid]));
                } else {
                  res.statusCode = 404;
                  res.end(JSON.stringify({ error: 'Record not found' }));
                }
              } catch (e: any) {
                res.statusCode = 400;
                res.end(JSON.stringify({ error: e.message }));
              }
            });
            break;
          }

          case 'DELETE': {
            if (!recordDcid) {
              res.statusCode = 400;
              res.end(JSON.stringify({ error: 'DCID required for delete' }));
              break;
            }
            if (store[recordDcid]) {
              delete store[recordDcid];
              res.statusCode = 204;
              res.end();
            } else {
              res.statusCode = 404;
              res.end(JSON.stringify({ error: 'Record not found' }));
            }
            break;
          }

          default:
            next();
        }
      });
    },
  };
}
