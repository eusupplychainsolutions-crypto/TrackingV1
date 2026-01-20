const express = require('express');
const session = require('express-session');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;
import { pca } from "./auth/msalClient.js";

// CORS
app.use(cors({ origin: true, credentials: true }));

// Body parser
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Session (Render supports this)
app.use(
  session({
    secret: 'secret123',
    resave: false,
    saveUninitialized: false,
    cookie: {
      maxAge: 1000 * 60 * 60,
      sameSite: 'lax'
    }
  })
);

const users = [
  {
    id: 1,
    name: 'Test Customer',
    email: 'test@demo.com',
    password: '123456'
  }
];

const shipments = [
  {
    id: 1,
    customerId: 1,
    containerNo: 'CMAU1234567',
    vessel: 'YM WELLNESS',
    voyage: '123W',
    pol: 'BRISBANE',
    pod: 'SHANGHAI',
    etd: '2025-12-15',
    eta: '2025-12-28',
    status: 'On board'
  }
];

function requireLogin(req, res, next) {
  if (!req.session.user) return res.status(401).json({ message: 'Not logged in' });
  next();
}

app.post('/api/login', (req, res) => {
  const { email, password } = req.body;

  const user = users.find(
    (u) => u.email === email && u.password === password
  );

  if (!user) return res.status(401).json({ message: 'Invalid email or password' });

  req.session.user = {
    id: user.id,
    name: user.name,
    email: user.email
  };

  res.json({ message: 'Login success', user: req.session.user });
});

app.get('/api/me', requireLogin, (req, res) => {
  res.json(req.session.user);
});

app.get('/api/shipments', requireLogin, (req, res) => {
  const u = req.session.user;
  res.json(shipments.filter(s => s.customerId === u.id));
});

// Root path
app.get('/', (req, res) => {
  res.send('Cargo Tracking API is running OK.');
});
app.get("/api/_health/graph", async (req, res) => {
  try {
    const result = await pca.acquireTokenSilent({
      scopes: ["User.Read"],
      account: pca.getAllAccounts()[0],
    });

    res.json({
      ok: true,
      message: "Microsoft Graph reachable",
      tenantId: result.tenantId,
      expiresOn: result.expiresOn,
    });
  } catch (err) {
    res.status(500).json({
      ok: false,
      error: err.message,
      errorCode: err.errorCode,
    });
  }
});

// THIS IS REQUIRED FOR RENDER
app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server running on port ${PORT}`);
});
