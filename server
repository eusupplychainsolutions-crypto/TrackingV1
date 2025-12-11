const express = require('express');
const session = require('express-session');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

// 允许跨域（Wix 或其他前端域名访问 API）
app.use(
  cors({
    origin: true,
    credentials: true
  })
);

// 解析 JSON 和表单数据
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Session 用于保存登录状态
app.use(
  session({
    secret: 'your-secret-key',
    resave: false,
    saveUninitialized: false,
    cookie: {
      maxAge: 1000 * 60 * 60, // 1 小时
      sameSite: 'lax'
    }
  })
);

// ====== 模拟用户数据（后续可换数据库）======
const users = [
  {
    id: 1,
    name: 'Test Customer',
    email: 'test@demo.com',
    password: '123456'
  }
];

// ====== 模拟船期数据 ======
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
  },
  {
    id: 2,
    customerId: 1,
    containerNo: 'MSCU7654321',
    vessel: 'MSC AURORA',
    voyage: '045E',
    pol: 'SYDNEY',
    pod: 'NINGBO',
    etd: '2025-12-20',
    eta: '2026-01-03',
    status: 'Booked'
  }
];

// ====== 中间件：检查是否已登录 ======
function requireLogin(req, res, next) {
  if (!req.session.user) {
    return res.status(401).json({ message: 'Not logged in' });
  }
  next();
}

// ====== 登录接口 ======
app.post('/api/login', (req, res) => {
  const { email, password } = req.body;

  const user = users.find(
    (u) => u.email === email && u.password === password
  );

  if (!user) {
    return res.status(401).json({ message: 'Invalid email or password' });
  }

  req.session.user = {
    id: user.id,
    name: user.name,
    email: user.email
  };

  res.json({ message: 'Login success', user: req.session.user });
});

// ====== 当前登录用户 ======
app.get('/api/me', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ message: 'Not logged in' });
  }
  res.json(req.session.user);
});

// ====== 登出 ======
app.post('/api/logout', (req, res) => {
  req.session.destroy(() => {
    res.json({ message: 'Logged out' });
  });
});

// ====== 查询集装箱 ======
app.get('/api/shipments', requireLogin, (req, res) => {
  const user = req.session.user;
  const { containerNo, vessel } = req.query;

  let result = shipments.filter((s) => s.customerId === user.id);

  if (containerNo) {
    result = result.filter((s) =>
      s.containerNo.toUpperCase().includes(containerNo.toUpperCase())
    );
  }

  if (vessel) {
    result = result.filter((s) =>
      s.vessel.toUpperCase().includes(vessel.toUpperCase())
    );
  }

  res.json(result);
});

// 基础测试路径
app.get('/', (req, res) => {
  res.send('Cargo Tracking API is running.');
});

// 启动服务器
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
