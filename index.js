import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import webpush from "web-push";
import cron from "node-cron";
import { createClient } from "@supabase/supabase-js";

dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

const supabase = createClient(
    process.env.SUPABASE_URL,
    process.env.SUPABASE_SERVICE_ROLE
);

webpush.setVapidDetails(
  "mailto:your@email.com",
  process.env.VAPID_PUBLIC_KEY,
  process.env.VAPID_PRIVATE_KEY
);

// ✅ CREATE USER (ANTI 429)
app.post("/create-user", verifyAdmin, async (req, res) => {

    const { email, password, name, role } = req.body;

    try {

        const { data, error } = await supabase.auth.admin.createUser({
            email,
            password,
            email_confirm: true
        });

        if (error) throw error;

        // insert ke profiles
        await supabase.from("profiles").insert({
            id: data.user.id,
            name,
            role
        });

        res.json({ success: true });

    } catch (err) {
        console.error(err);
        res.status(400).json({ error: err.message });
    }
});

app.delete("/delete-user/:id", verifyAdmin, async (req, res) => {

    const userId = req.params.id;

    try {

        // hapus attendance dulu
        const { error: attendanceError } = await supabase
            .from("attendance")
            .delete()
            .eq("userid", userId);

        if (attendanceError) throw attendanceError;

        // ❗ hapus dari profiles (optional tapi disarankan)
        const { error: profileError } = await supabase
            .from("profiles")
            .delete()
            .eq("id", userId);

        if (profileError) throw profileError;

        // ❗ hapus dari auth
        const { error } = await supabase.auth.admin.deleteUser(userId);

        if (error) throw error;

        res.json({ success: true });

    } catch (err) {
        console.error(err);
        res.status(400).json({ error: err.message });
    }
});

app.post("/add-attendance", verifyAdmin, async (req, res) => {
  try {
    const { userid, latitude, longitude, status, reason } = req.body;

    if (!userid || !status) {
      return res.status(400).json({ error: "User dan status wajib" });
    }

    if (status === "Izin" && !reason) {
      return res.status(400).json({ error: "Alasan wajib untuk izin" });
    }

    const { error } = await supabase
      .from("attendance")
      .insert([{
        userid: userid,
        latitude: latitude !== undefined ? latitude : -6.2,   // optional (atau kirim dari frontend)
        longitude: longitude !== undefined ? longitude : 106.8, // optional
        status: status,
        reason: status === "Izin" ? reason : null
      }]);

    if (error) throw error;

    res.json({ message: "Attendance berhasil ditambahkan" });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

import ExcelJS from "exceljs";

app.get("/export-attendance-rekap", verifyAdmin, async (req, res) => {
  try {

    // 🔥 ambil data
    const { data, error } = await supabase
        .from("attendance_view")
        .select("*")
        .order("tanggal", { ascending: true });

    if (error) throw error;

    // =========================
    // 🧠 1. GROUP TANGGAL PER BULAN
    // =========================
    const monthMap = {}; // { "Januari 2026": [1,2,3...] }

    data.forEach(item => {
      const d = new Date(item.tanggal);
      const monthName = d.toLocaleDateString("id-ID", { month: "long", year: "numeric" });
      const day = d.getDate();

      if (!monthMap[monthName]) {
        monthMap[monthName] = new Set();
      }

      monthMap[monthName].add(day);
    });

    // sort tanggal
    for (let m in monthMap) {
      monthMap[m] = Array.from(monthMap[m]).sort((a, b) => a - b);
    }

    // =========================
    // 🧠 2. GROUP USER
    // =========================
    const users = {};

    data.forEach(item => {
      const name = item.name || "Unknown";
      const d = new Date(item.tanggal);
      const key = `${d.toLocaleDateString("id-ID", { month: "long", year: "numeric" })}-${d.getDate()}`;

      if (!users[name]) users[name] = {};
      users[name][key] = item.status.toUpperCase();
    });

    // =========================
    // 📊 3. BUAT EXCEL
    // =========================
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Rekap Absensi");

    // =========================
    // 🟦 JUDUL
    // =========================
    ws.mergeCells("A1:B1");
    ws.getCell("A1").value = "REKAP ABSENSI SMB 2026";
    ws.getCell("A1").font = { bold: true, size: 14 };

    // =========================
    // 🟦 HEADER ROW 1 (BULAN)
    // =========================
    let colIndex = 3; // mulai dari kolom ke-3

    ws.getCell("A2").value = "No";
    ws.getCell("B2").value = "Nama";

    ws.mergeCells("A2:A3");
    ws.mergeCells("B2:B3");

    for (let month in monthMap) {
      const days = monthMap[month];
      const startCol = colIndex;
      const endCol = colIndex + days.length - 1;

      ws.mergeCells(2, startCol, 2, endCol);
      ws.getCell(2, startCol).value = month;

      days.forEach(day => {
        ws.getCell(3, colIndex).value = day;
        colIndex++;
      });
    }

    // HEADER TOTAL
    const totalStartCol = colIndex;

    ws.mergeCells(2, totalStartCol, 2, totalStartCol + 2);

    ws.getCell(2, totalStartCol).value = "TOTAL";

    ws.getCell(3, totalStartCol).value = "HADIR";
    ws.getCell(3, totalStartCol + 1).value = "IZIN";
    ws.getCell(3, totalStartCol + 2).value = "TERLAMBAT";

    // =========================
    // 🎨 STYLE HEADER
    // =========================
    [2, 3].forEach(rowNum => {
      ws.getRow(rowNum).eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: rowNum === 2 ? "9DC3E6" : "BDD7EE" }
        };
      });
    });

    // =========================
    // 👤 4. ISI DATA
    // =========================
    let rowNum = 4;
    let no = 1;

    for (const name in users) {
      let hadir = 0;
      let izin = 0;
      let terlambat = 0;
      let col = 3;  

      const cellNo = ws.getCell(rowNum, 1);
      cellNo.value = no++;
      const cellName = ws.getCell(rowNum, 2);
      cellName.value = name;

      cellNo.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };

      cellName.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };

      for (let month in monthMap) {
        monthMap[month].forEach(day => {
          const key = `${month}-${day}`;
          const status = users[name][key] || "";

          const cell = ws.getCell(rowNum, col);
          cell.value = status;

          if (status === "HADIR") hadir++;
          if (status === "IZIN") izin++;
          if (status === "TERLAMBAT") terlambat++;

          // 🎨 warna
          if (status === "HADIR") {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "C6EFCE" } };
          } else if (status === "IZIN") {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "F4B084" } };
          } else if (status === "TERLAMBAT") {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEB9C" } };
          }

          cell.alignment = { horizontal: "center" };
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
          };

          col++;
        });
      }

      const cellHadir = ws.getCell(rowNum, col);
      cellHadir.value = hadir;
      const cellIzin = ws.getCell(rowNum, col + 1);
      cellIzin.value = izin;
      const cellTerlambat = ws.getCell(rowNum, col + 2);
      cellTerlambat.value = terlambat;

      cellHadir.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "C6EFCE" } };
      cellIzin.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "F4B084" } };
      cellTerlambat.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEB9C" } };


      // style
      for (let i = 0; i < 3; i++) {
        const c = ws.getCell(rowNum, col + i);
        c.alignment = { horizontal: "center" };
        c.font = { bold: true };
        c.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" }
        };
      }

      rowNum++;
    }

    // =========================
    // 📏 AUTO WIDTH
    // =========================
    ws.getColumn(1).width = 5;
    ws.getColumn(2).width = 25;

    for (let i = 3; i <= ws.columnCount; i++) {
      ws.getColumn(i).width = 10;
    }

    ws.getColumn(colIndex).width = 12;
    ws.getColumn(colIndex + 1).width = 12;
    ws.getColumn(colIndex + 2).width = 15;

    // =========================
    // 📥 DOWNLOAD
    // =========================
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.setHeader(
      "Content-Disposition",
      "attachment; filename=rekap-absensi-full.xlsx"
    );

    await wb.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get("/test-notif", async (req, res) => {
  try {
    await sendReminder(0, "ini test notif 🚀");
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/send-notif", async (req, res) => {
  try {
    const { alluser, message} = req.body;

    if (!alluser || !message) {
      return res.status(400).json({ error: "AllUser dan message wajib" });
    }

    let alluserValue = 1;
    if (alluser === "0") {
      alluserValue = 0;
    }

    await sendReminder(alluserValue, message);
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get("/ping", (_, res) => {
  console.log("Received ping at", new Date().toISOString());
  res.send("ok");
});

app.post("/cron-reminder", async (req, res) => {
    try {
        const token = req.headers["CRON_SECRET"];
        const { alluser, message } = req.body;

        if (token !== process.env.CRON_SECRET) {
            console.warn("Unauthorized cron access attempt with token:", token);
            return res.status(403).json({ error: "Unauthorized" });
        }

        if (!alluser || !message) {
          return res.status(400).json({ error: "AllUser dan message wajib" });
        }

        console.log("External cron triggered, alluser:", alluser, "message:", message);

        await sendReminder(0, message);

        res.json({ success: true });

    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

async function verifyAdmin(req, res, next) {

    const authHeader = req.headers.authorization;

    if (!authHeader) {
        return res.status(401).json({ error: "No token" });
    }

    const token = authHeader.replace("Bearer ", "");

    try {

        // ✅ verify user dari token
        const { data, error } = await supabase.auth.getUser(token);

        if (error || !data.user) {
            return res.status(401).json({ error: "Invalid token" });
        }

        const userId = data.user.id;

        // ✅ cek role di profiles
        const { data: profile } = await supabase
            .from("profiles")
            .select("role")
            .eq("id", userId)
            .single();

        if (!profile || profile.role !== "admin") {
            return res.status(403).json({ error: "Forbidden (admin only)" });
        }

        // simpan user ke request
        req.user = data.user;

        next();

    } catch (err) {
        res.status(500).json({ error: err.message });
    }
}

// ✅ KIRIM REMINDER PUSH NOTIFICATION
async function sendReminder(alluser, message) {
  console.log("All user:", alluser);
  console.log("Mengirim reminder:", message);
  const { start, end } = getTodayRange();

  // ambil yg sudah absen
  const { data: attended } = await supabase
    .from("attendance")
    .select("userid")
    .gte("timestamp", start.toISOString())
    .lte("timestamp", end.toISOString());

  const attendedIds = new Set(attended.map(a => a.userid));

  // ambil semua user
  const { data: users } = await supabase
    .from("profiles")
    .select("id, name");

  // filter belum absen
  const notYet = users.filter(u => !attendedIds.has(u.id));
  console.log("User yang belum absen:", notYet.map(u => u.name).join(", "));

  // ambil subscription aktif
  const { data: subs } = await supabase
    .from("push_subscriptions")
    .select("*")
    .eq("is_active", true);

  console.log("Push subscriptions aktif:", subs.length);

  const userMap = {};
  users.forEach(u => userMap[u.id] = u.name);

  for (const sub of subs) {

    // hanya kirim ke user yg belum absen

    if (alluser === 0 && !notYet.find(u => u.id === sub.user_id)) continue;

    const name = userMap[sub.user_id] || "User";
    console.log(`Mengirim notif ke ${name} (${sub.endpoint})`);

    const payload = JSON.stringify({
      title: "Reminder Absensi",
      body: `Halo ${name}, ${message}`
    });

    try {
      await webpush.sendNotification({
        endpoint: sub.endpoint,
        keys: {
          p256dh: sub.p256dh,
          auth: sub.auth
        }
      }, payload);
      console.log("SUCCESS SEND");

    } catch (err) {
      console.error("PUSH ERROR:", err);
      if (err.statusCode === 410) {
        await supabase
          .from("push_subscriptions")
          .delete()
          .eq("endpoint", sub.endpoint);
      }
    }
  }
}

// helper untuk dapatkan range waktu hari ini (untuk query attendance)
function getTodayRange() {
  const start = new Date();
  start.setHours(0,0,0,0);

  const end = new Date();
  end.setHours(23,59,59,999);

  return { start, end };
}


// Schedule cron job untuk kirim reminder setiap minggu
// minggu jam 18:00
cron.schedule("0 9 * * 0", async () => {
  console.log("Cron job running at 09:00 every Sunday");
  await sendReminder(0, "waktunya absen ya 👋");
  // await sendReminder(0, "[TESTING NOTIF PRE-PROD] waktunya absen ya 👋");
});

// minggu jam 09:55
cron.schedule("55 9 * * 0", async () => {
  await sendReminder(0, "yuk absen, sisa 5 menit lagi ⏰");
});

// minggu jam 11:50
cron.schedule("50 11 * * 0", async () => {
  await sendReminder(0, "absen akan ditutup dalam waktu 10 menit lagi 🚫");
});

// minggu jam 11:55
cron.schedule("55 11 * * 0", async () => {
  await sendReminder(0, "absen gasiihh, sisa 5 menit lagi tutup nih anjay 😑🚫");
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
    console.log("Server Running on port", PORT);
});