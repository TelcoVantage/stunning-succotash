// Minimal mock of the Genesys Cloud endpoints used by Get-GenesysPriorityCallReport.ps1
const http = require('http');
const url = require('url');

const DIV = 'div-1';
const Q1 = 'queue-1';
const S1 = 'skill-1';
const L1 = 'lang-1';
const A = 'user-a', B = 'user-b', C = 'user-c';

// Scenario timeline (UTC, 2026-09-02)
// c1: prio 0, no skill,  enters 10:00:00, answered by A 10:00:05, handled until 10:05:00
// c2: prio 5, skill S1,  enters 10:00:30, answered by B 10:02:00 (waited 90s), handled to 10:06:00
//     B is IDLE 10:00:00-10:01:00 (eligible idle 30s while c2 waits), then takes c3 (jump ahead!)
// c3: prio 0, no skill,  enters 10:00:40, answered by B 10:01:00, handled until 10:01:50
// c4: prio 0, no skill,  enters 10:03:00, abandoned 10:03:20 ; C (no skill) idle whole time -> eligible idle
// c5: prio 0, no skill,  enters 10:10:00, answered by A 10:10:02 ; conversation GET returns 404 (purged)
// c6: prio 100, no skill, enters 10:20:00, answered by C 10:20:40  (older, lower priority)
// c7: prio 400, no skill, enters 10:20:10, answered by A 10:20:20  -> overtakes c6 = PRIORITY HONOURED
const T = (hms) => `2026-09-02T${hms}.000Z`;

function seg(type, start, end, extra) { return Object.assign({ segmentType: type, segmentStart: T(start), segmentEnd: T(end) }, extra || {}); }

function conv(id, start, end, acdSeg, acdSession, agentParts) {
  return {
    conversationId: id, conversationStart: T(start), conversationEnd: T(end), originatingDirection: 'inbound', divisionIds: [DIV],
    participants: [
      { participantId: id + '-cust', purpose: 'customer', sessions: [{ mediaType: 'voice', direction: 'inbound', ani: 'tel:+4412345', dnis: 'tel:+4499999', segments: [seg('interact', start, end)] }] },
      { participantId: id + '-acd', purpose: 'acd', sessions: [Object.assign({ mediaType: 'voice', direction: 'inbound', segments: [acdSeg] }, acdSession)] },
      ...agentParts
    ]
  };
}
function agent(uid, alertStart, answer, end) {
  return { participantId: uid + '-p', purpose: 'agent', userId: uid, sessions: [{ mediaType: 'voice', direction: 'inbound', segments: [seg('alert', alertStart, answer), seg('interact', answer, end), seg('wrapup', end, end)] }] };
}

const conversations = [
  conv('c1', '10:00:00', '10:05:00', seg('interact', '10:00:00', '10:00:05', { queueId: Q1, disconnectType: 'transfer' }), { usedRouting: 'Standard', requestedRoutings: ['Standard'] }, [agent(A, '10:00:02', '10:00:05', '10:05:00')]),
  conv('c2', '10:00:30', '10:06:00', seg('interact', '10:00:30', '10:02:00', { queueId: Q1, disconnectType: 'transfer' }), { usedRouting: 'Standard', requestedRoutings: ['Standard'], requestedRoutingSkillIds: [S1] }, [agent(B, '10:01:57', '10:02:00', '10:06:00')]),
  conv('c3', '10:00:40', '10:01:50', seg('interact', '10:00:40', '10:01:00', { queueId: Q1, disconnectType: 'transfer' }), { usedRouting: 'Standard', requestedRoutings: ['Standard'] }, [agent(B, '10:00:58', '10:01:00', '10:01:50')]),
  conv('c4', '10:03:00', '10:03:20', seg('interact', '10:03:00', '10:03:20', { queueId: Q1, disconnectType: 'client' }), { usedRouting: 'Standard', requestedRoutings: ['Standard'] }, []),
  conv('c5', '10:10:00', '10:12:00', seg('interact', '10:10:00', '10:10:02', { queueId: Q1, disconnectType: 'transfer' }), { usedRouting: 'Standard', requestedRoutings: ['Standard'] }, [agent(A, '10:10:00', '10:10:02', '10:12:00')]),
  conv('c6', '10:20:00', '10:24:00', seg('interact', '10:20:00', '10:20:40', { queueId: Q1, disconnectType: 'transfer' }), { usedRouting: 'Standard', requestedRoutings: ['Standard'] }, [agent(C, '10:20:38', '10:20:40', '10:24:00')]),
  conv('c7', '10:20:10', '10:25:00', seg('interact', '10:20:10', '10:20:20', { queueId: Q1, disconnectType: 'transfer' }), { usedRouting: 'Standard', requestedRoutings: ['Standard'] }, [agent(A, '10:20:18', '10:20:20', '10:25:00')]),
];
const priorities = { c1: 0, c2: 5, c3: 0, c4: 0, c6: 100, c7: 400 }; // c5 -> 404

function convDetail(id) {
  if (!(id in priorities)) return null;
  const crd = { queue: { id: Q1 }, priority: priorities[id], skills: id === 'c2' ? [{ id: S1, name: 'VIP Skill' }] : [], scoredAgents: [] };
  return { id, participants: [
    { id: 'p1', purpose: 'customer' },
    { id: 'p2', purpose: 'acd', queueId: Q1, connectedTime: conversations.find(c => c.conversationId === id).conversationStart, conversationRoutingData: crd }
  ] };
}

const users = {
  [A]: { id: A, name: 'Alice Agent', skills: [{ id: S1, name: 'VIP Skill' }], languages: [] },
  [B]: { id: B, name: 'Bob Agent', skills: [{ id: S1, name: 'VIP Skill' }], languages: [] },
  [C]: { id: C, name: 'Carol Agent', skills: [], languages: [] },
};
function rs(status, start, end) { return { routingStatus: status, startTime: T(start), endTime: end ? T(end) : undefined }; }
const routing = {
  [A]: [rs('IDLE', '09:00:00', '10:00:02'), rs('INTERACTING', '10:00:02', '10:05:00'), rs('IDLE', '10:05:00', '10:10:00'), rs('INTERACTING', '10:10:00', '10:20:18'), rs('IDLE', '10:20:18', '10:20:19'), rs('INTERACTING', '10:20:19', '10:25:00'), rs('OFF_QUEUE', '10:25:00', null)],
  [B]: [rs('IDLE', '10:00:00', '10:00:58'), rs('INTERACTING', '10:00:58', '10:01:50'), rs('IDLE', '10:01:50', '10:01:57'), rs('INTERACTING', '10:01:57', '10:06:00'), rs('OFF_QUEUE', '10:06:00', null)],
  [C]: [rs('IDLE', '09:30:00', '10:15:00'), rs('INTERACTING', '10:15:00', '10:20:37'), rs('IDLE', '10:20:37', '10:20:39'), rs('INTERACTING', '10:20:39', '10:30:00'), rs('OFF_QUEUE', '10:30:00', null)],
};

let hits = {};
const server = http.createServer((req, res) => {
  let body = '';
  req.on('data', d => body += d);
  req.on('end', () => {
    const u = url.parse(req.url, true);
    const p = u.pathname;
    hits[req.method + ' ' + p] = (hits[req.method + ' ' + p] || 0) + 1;
    const send = (code, obj) => { res.writeHead(code, { 'Content-Type': 'application/json' }); res.end(JSON.stringify(obj)); };
    const auth = req.headers['authorization'] || '';
    if (p === '/oauth/token') {
      // expect Basic base64('client-id:client-secret')
      if (auth !== 'Basic ' + Buffer.from('client-id:client-secret').toString('base64')) return send(401, { error: 'bad basic header: ' + auth });
      return send(200, { access_token: 'tok-123', token_type: 'bearer' });
    }
    if (auth !== 'Bearer tok-123') return send(401, { message: 'no bearer' });
    if (p === '/api/v2/authorization/divisions') return send(200, { entities: [{ id: 'div-0', name: 'Home' }, { id: DIV, name: 'Customer Service' }], pageCount: 1 });
    if (p === '/api/v2/routing/skills') return send(200, { entities: [{ id: S1, name: 'VIP Skill' }], pageCount: 1 });
    if (p === '/api/v2/routing/languages') return send(200, { entities: [{ id: L1, name: 'English' }], pageCount: 1 });
    if (p === '/api/v2/analytics/conversations/details/query') {
      const b = JSON.parse(body);
      if (b.paging.pageNumber > 1) return send(200, { conversations: [], totalHits: conversations.length });
      const divOk = JSON.stringify(b.conversationFilters).includes(DIV);
      if (!divOk) return send(400, { message: 'missing division filter' });
      return send(200, { conversations, totalHits: conversations.length });
    }
    if (p === '/api/v2/routing/queues/' + Q1) return send(200, { id: Q1, name: 'Customer Service Line' });
    if (p === '/api/v2/routing/queues/' + Q1 + '/members') return send(200, { entities: [A, B, C].map(id => ({ id, joined: true, user: users[id] })), pageCount: 1 });
    if (p.startsWith('/api/v2/conversations/')) {
      const d = convDetail(p.split('/').pop());
      if (!d) return send(404, { message: 'not found' });
      return send(200, d);
    }
    if (p === '/api/v2/users') {
      const ids = [].concat(u.query.id || []);
      if (u.query.expand !== 'skills,languages') return send(400, { message: 'expand' });
      return send(200, { entities: ids.filter(i => users[i]).map(i => users[i]) });
    }
    if (p === '/api/v2/analytics/users/details/query') {
      const b = JSON.parse(body);
      const ids = b.userFilters[0].predicates.map(x => x.value);
      if (b.paging.pageNumber > 1) return send(200, { userDetails: [] });
      return send(200, { userDetails: ids.map(i => ({ userId: i, routingStatus: routing[i] || [], primaryPresence: [] })) });
    }
    if (p === '/__hits') return send(200, hits);
    send(404, { message: 'unmocked ' + req.method + ' ' + p });
  });
});
server.listen(8091, () => console.log('mock listening on 8091'));
