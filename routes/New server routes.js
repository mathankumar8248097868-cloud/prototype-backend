// ================= SESSION CHECK ENDPOINTS =================
// Add these routes to server.js BEFORE the static files middleware

// Check if regular user is logged in (called by report.html auth guard)
app.get("/api/auth/check", (req, res) => {
  res.json({ loggedIn: !!req.session.user });
});

// Check if admin is logged in (called by adminpage.html auth guard)
app.get("/api/auth/admin/check", (req, res) => {
  res.json({ loggedIn: !!req.session.admin });
});

// User logout
app.post("/api/logout", (req, res) => {
  req.session.destroy();
  res.json({ success: true });
});

// Admin logout
app.post("/api/admin/logout", (req, res) => {
  req.session.destroy();
  res.json({ success: true });
});