import { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  BarChart,
  Bar,
  LineChart,
  Line,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  AreaChart,
  Area,
} from "recharts";

// ── 기본 샘플 데이터 ──────────────────────────────────────
const sampleHeadcount = [
  { month: "1월", total: 342, join: 18, leave: 8 },
  { month: "2월", total: 352, join: 15, leave: 5 },
  { month: "3월", total: 361, join: 20, leave: 11 },
  { month: "4월", total: 370, join: 14, leave: 5 },
  { month: "5월", total: 378, join: 12, leave: 4 },
  { month: "6월", total: 384, join: 9, leave: 3 },
];
const sampleDept = [
  { name: "개발", value: 120, color: "#00D4AA" },
  { name: "영업", value: 85, color: "#FF6B6B" },
  { name: "마케팅", value: 62, color: "#FFD166" },
  { name: "HR", value: 38, color: "#74B9FF" },
  { name: "재무", value: 45, color: "#A29BFE" },
  { name: "운영", value: 34, color: "#FD79A8" },
];
const samplePipeline = [
  {
    position: "시니어 개발자",
    total: 3,
    s1: 3,
    s2: 2,
    s3: 1,
    s4: 0,
    status: "진행중",
    urgent: true,
  },
  {
    position: "마케팅 매니저",
    total: 1,
    s1: 1,
    s2: 1,
    s3: 1,
    s4: 1,
    status: "최종단계",
    urgent: false,
  },
  {
    position: "재무 분석가",
    total: 2,
    s1: 2,
    s2: 2,
    s3: 0,
    s4: 0,
    status: "진행중",
    urgent: false,
  },
  {
    position: "영업 팀장",
    total: 1,
    s1: 1,
    s2: 0,
    s3: 0,
    s4: 0,
    status: "초기단계",
    urgent: true,
  },
  {
    position: "UX 디자이너",
    total: 2,
    s1: 2,
    s2: 2,
    s3: 2,
    s4: 0,
    status: "심사중",
    urgent: false,
  },
];
const sampleBudget = [
  { month: "1월", budget: 4200, actual: 4050 },
  { month: "2월", budget: 4200, actual: 4100 },
  { month: "3월", budget: 4300, actual: 4280 },
  { month: "4월", budget: 4300, actual: 4350 },
  { month: "5월", budget: 4500, actual: 4420 },
  { month: "6월", budget: 4500, actual: 4510 },
];
const sampleBreakdown = [
  { label: "기본급", budget: 3200, actual: 3180, color: "#00D4AA" },
  { label: "성과급", budget: 600, actual: 650, color: "#FFD166" },
  { label: "복리후생", budget: 400, actual: 390, color: "#74B9FF" },
  { label: "교육훈련", budget: 300, actual: 290, color: "#A29BFE" },
];

const DEPT_COLORS = [
  "#00D4AA",
  "#FF6B6B",
  "#FFD166",
  "#74B9FF",
  "#A29BFE",
  "#FD79A8",
  "#55EFC4",
  "#FDCB6E",
];
const stageColors = ["#74B9FF", "#A29BFE", "#FFD166", "#00D4AA"];
const statusColors: Record<string, string> = {
  진행중: "#74B9FF",
  최종단계: "#00D4AA",
  초기단계: "#8B949E",
  심사중: "#FFD166",
};
const tabs = [
  { id: "overview", label: "개요" },
  { id: "headcount", label: "인원 현황" },
  { id: "pipeline", label: "채용 파이프라인" },
  { id: "budget", label: "인건비 현황" },
];

// ── 엑셀 시트 → 객체 배열 변환 ───────────────────────────
function sheetToJson(wb: XLSX.WorkBook, name: string) {
  const ws = wb.Sheets[name];
  if (!ws) return null;
  return XLSX.utils.sheet_to_json(ws) as Record<string, any>[];
}

// ── 메인 컴포넌트 ─────────────────────────────────────────
export default function HRDashboard() {
  const [activeTab, setActiveTab] = useState("overview");
  const [headcount, setHeadcount] = useState(sampleHeadcount);
  const [deptData, setDeptData] = useState(sampleDept);
  const [pipeline, setPipeline] = useState(samplePipeline);
  const [budget, setBudget] = useState(sampleBudget);
  const [breakdown, setBreakdown] = useState(sampleBreakdown);
  const [uploadedSheets, setUploadedSheets] = useState<string[]>([]);
  const [dragging, setDragging] = useState(false);
  const [toast, setToast] = useState("");

  const showToast = (msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(""), 3000);
  };

  const processFile = useCallback((file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });
        const loaded: string[] = [];

        // ── 인원현황 시트 ─────────────────────────────────
        const hc = sheetToJson(wb, "인원현황");
        if (hc) {
          setHeadcount(
            hc.map((r) => ({
              month: String(r["월"] ?? r["month"] ?? ""),
              total: Number(r["직원수"] ?? r["total"] ?? 0),
              join: Number(r["입사"] ?? r["join"] ?? 0),
              leave: Number(r["퇴사"] ?? r["leave"] ?? 0),
            }))
          );
          loaded.push("인원현황");
        }

        // ── 부서현황 시트 ─────────────────────────────────
        const dp = sheetToJson(wb, "부서현황");
        if (dp) {
          setDeptData(
            dp.map((r, i) => ({
              name: String(r["부서"] ?? r["name"] ?? ""),
              value: Number(r["인원"] ?? r["value"] ?? 0),
              color: DEPT_COLORS[i % DEPT_COLORS.length],
            }))
          );
          loaded.push("부서현황");
        }

        // ── 채용파이프라인 시트 ───────────────────────────
        const pl = sheetToJson(wb, "채용파이프라인");
        if (pl) {
          setPipeline(
            pl.map((r) => ({
              position: String(r["포지션"] ?? r["position"] ?? ""),
              total: Number(r["채용목표"] ?? r["total"] ?? 0),
              s1: Number(r["서류"] ?? r["s1"] ?? 0),
              s2: Number(r["1차면접"] ?? r["s2"] ?? 0),
              s3: Number(r["2차면접"] ?? r["s3"] ?? 0),
              s4: Number(r["최종"] ?? r["s4"] ?? 0),
              status: String(r["상태"] ?? r["status"] ?? "진행중"),
              urgent:
                String(r["긴급"] ?? r["urgent"] ?? "").toLowerCase() ===
                  "true" || r["긴급"] === 1,
            }))
          );
          loaded.push("채용파이프라인");
        }

        // ── 인건비 시트 ───────────────────────────────────
        const bd = sheetToJson(wb, "인건비");
        if (bd) {
          setBudget(
            bd.map((r) => ({
              month: String(r["월"] ?? r["month"] ?? ""),
              budget: Number(r["예산"] ?? r["budget"] ?? 0),
              actual: Number(r["실적"] ?? r["actual"] ?? 0),
            }))
          );
          loaded.push("인건비");
        }

        // ── 인건비항목 시트 ───────────────────────────────
        const bk = sheetToJson(wb, "인건비항목");
        if (bk) {
          setBreakdown(
            bk.map((r, i) => ({
              label: String(r["항목"] ?? r["label"] ?? ""),
              budget: Number(r["예산"] ?? r["budget"] ?? 0),
              actual: Number(r["실적"] ?? r["actual"] ?? 0),
              color: DEPT_COLORS[i % DEPT_COLORS.length],
            }))
          );
          loaded.push("인건비항목");
        }

        if (loaded.length === 0) {
          showToast(
            "❌ 시트 이름을 확인해주세요 (인원현황·부서현황·채용파이프라인·인건비·인건비항목)"
          );
        } else {
          setUploadedSheets(loaded);
          showToast(`✅ ${loaded.join(", ")} 시트 반영 완료!`);
        }
      } catch {
        showToast("❌ 파일을 읽는 중 오류가 발생했어요.");
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };
  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) processFile(file);
  };

  // ── KPI 계산 ─────────────────────────────────────────
  const lastRow = headcount[headcount.length - 1] ?? {
    total: 0,
    join: 0,
    leave: 0,
  };
  const prevRow = headcount[headcount.length - 2] ?? { total: 0 };
  const totalEmp = lastRow.total;
  const diffEmp = totalEmp - prevRow.total;
  const turnover =
    totalEmp > 0 ? ((lastRow.leave / totalEmp) * 100).toFixed(1) : "0.0";
  const lastBudget = budget[budget.length - 1] ?? { budget: 1, actual: 0 };
  const budgetRate = ((lastBudget.actual / lastBudget.budget) * 100).toFixed(1);
  const urgentCount = pipeline.filter((d) => d.urgent).length;
  const totalOpenings = pipeline.reduce((s, d) => s + d.total, 0);

  const kpis = [
    {
      label: "총 직원수",
      value: `${totalEmp}명`,
      sub: `전월 대비 ${diffEmp >= 0 ? "+" : ""}${diffEmp}`,
      icon: "👥",
      color: "#00D4AA",
    },
    {
      label: "이번달 신규 입사",
      value: `${lastRow.join}명`,
      sub: "목표 대비 75%",
      icon: "🟢",
      color: "#74B9FF",
    },
    {
      label: "이직률",
      value: `${turnover}%`,
      sub: "업계 평균 3.8%",
      icon: "📉",
      color: "#A29BFE",
    },
    {
      label: "예산 집행률",
      value: `${budgetRate}%`,
      sub: Number(budgetRate) > 100 ? "⚠️ 초과" : "✅ 정상 범위",
      icon: "📊",
      color: Number(budgetRate) > 100 ? "#FF6B6B" : "#FFD166",
    },
  ];

  return (
    <div
      style={{
        fontFamily: "'Noto Sans KR', sans-serif",
        background: "#0D1117",
        minHeight: "100vh",
        color: "#E6EDF3",
      }}
    >
      {/* Toast */}
      {toast && (
        <div
          style={{
            position: "fixed",
            bottom: 28,
            left: "50%",
            transform: "translateX(-50%)",
            background: "#1C2128",
            border: "1px solid #30363D",
            borderRadius: 10,
            padding: "12px 24px",
            fontSize: 13,
            fontWeight: 600,
            zIndex: 9999,
            boxShadow: "0 8px 32px #0008",
          }}
        >
          {toast}
        </div>
      )}

      {/* Nav */}
      <div
        style={{
          background: "rgba(22,27,34,0.97)",
          borderBottom: "1px solid #21262D",
          padding: "0 32px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          height: 60,
          position: "sticky",
          top: 0,
          zIndex: 100,
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div
            style={{
              width: 32,
              height: 32,
              borderRadius: 8,
              background: "linear-gradient(135deg,#00D4AA,#74B9FF)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontWeight: 800,
              fontSize: 15,
              color: "#0D1117",
            }}
          >
            H
          </div>
          <span style={{ fontWeight: 700, fontSize: 16 }}>HR 인사이트</span>
          <span
            style={{
              background: "#21262D",
              borderRadius: 20,
              padding: "2px 10px",
              fontSize: 11,
              color: "#8B949E",
            }}
          >
            2025 · 2분기
          </span>
        </div>
        <div style={{ display: "flex", gap: 4 }}>
          {tabs.map((t) => (
            <button
              key={t.id}
              onClick={() => setActiveTab(t.id)}
              style={{
                padding: "6px 14px",
                borderRadius: 6,
                border: "none",
                cursor: "pointer",
                fontSize: 12,
                fontWeight: 600,
                background: activeTab === t.id ? "#00D4AA22" : "transparent",
                color: activeTab === t.id ? "#00D4AA" : "#8B949E",
              }}
            >
              {t.label}
            </button>
          ))}
        </div>
        {/* 파일 업로드 버튼 */}
        <label
          style={{
            display: "flex",
            alignItems: "center",
            gap: 8,
            background: "#00D4AA",
            color: "#0D1117",
            borderRadius: 8,
            padding: "7px 16px",
            fontSize: 12,
            fontWeight: 700,
            cursor: "pointer",
          }}
        >
          📂 엑셀 업로드
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={onFileChange}
            style={{ display: "none" }}
          />
        </label>
      </div>

      <div style={{ padding: "28px 32px", maxWidth: 1400, margin: "0 auto" }}>
        {/* 드래그 앤 드롭 영역 */}
        <div
          onDragOver={(e) => {
            e.preventDefault();
            setDragging(true);
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={onDrop}
          style={{
            border: `2px dashed ${dragging ? "#00D4AA" : "#30363D"}`,
            borderRadius: 12,
            padding: "18px 24px",
            marginBottom: 20,
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            background: dragging ? "#00D4AA0A" : "transparent",
            transition: "all 0.2s",
          }}
        >
          <div>
            <div
              style={{
                fontWeight: 700,
                fontSize: 13,
                color: dragging ? "#00D4AA" : "#8B949E",
              }}
            >
              {dragging
                ? "파일을 여기에 놓으세요!"
                : "📎 엑셀 파일을 여기에 드래그하거나 우측 버튼으로 업로드하세요"}
            </div>
            <div style={{ fontSize: 11, color: "#6E7681", marginTop: 4 }}>
              시트명: 인원현황 · 부서현황 · 채용파이프라인 · 인건비 · 인건비항목
            </div>
          </div>
          {uploadedSheets.length > 0 && (
            <div style={{ display: "flex", gap: 6 }}>
              {uploadedSheets.map((s) => (
                <span
                  key={s}
                  style={{
                    fontSize: 10,
                    background: "#00D4AA22",
                    color: "#00D4AA",
                    borderRadius: 6,
                    padding: "3px 8px",
                    fontWeight: 700,
                  }}
                >
                  ✓ {s}
                </span>
              ))}
            </div>
          )}
        </div>

        {/* KPI */}
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(4,1fr)",
            gap: 16,
            marginBottom: 24,
          }}
        >
          {kpis.map((k, i) => (
            <div
              key={i}
              style={{
                background: "#161B22",
                border: "1px solid #21262D",
                borderRadius: 12,
                padding: "20px 22px",
                position: "relative",
                overflow: "hidden",
              }}
            >
              <div
                style={{
                  position: "absolute",
                  top: 0,
                  right: 0,
                  width: 80,
                  height: 80,
                  borderRadius: "0 12px 0 80px",
                  background: k.color + "0E",
                }}
              />
              <div style={{ fontSize: 22, marginBottom: 8 }}>{k.icon}</div>
              <div style={{ fontSize: 26, fontWeight: 800, color: k.color }}>
                {k.value}
              </div>
              <div style={{ fontSize: 12, color: "#8B949E", marginTop: 4 }}>
                {k.label}
              </div>
              <div
                style={{
                  fontSize: 11,
                  color: "#6E7681",
                  marginTop: 6,
                  borderTop: "1px solid #21262D",
                  paddingTop: 8,
                }}
              >
                {k.sub}
              </div>
            </div>
          ))}
        </div>

        {/* ── OVERVIEW ───────────────────────────────── */}
        {activeTab === "overview" && (
          <>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "2fr 1fr",
                gap: 16,
                marginBottom: 16,
              }}
            >
              <div
                style={{
                  background: "#161B22",
                  border: "1px solid #21262D",
                  borderRadius: 12,
                  padding: "22px 24px",
                }}
              >
                <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>
                  인원 변동 추이
                </div>
                <div
                  style={{ color: "#8B949E", fontSize: 12, marginBottom: 20 }}
                >
                  월별 총 직원수 · 입퇴사 현황
                </div>
                <ResponsiveContainer width="100%" height={200}>
                  <AreaChart data={headcount}>
                    <defs>
                      <linearGradient id="g1" x1="0" y1="0" x2="0" y2="1">
                        <stop
                          offset="5%"
                          stopColor="#74B9FF"
                          stopOpacity={0.25}
                        />
                        <stop
                          offset="95%"
                          stopColor="#74B9FF"
                          stopOpacity={0}
                        />
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke="#21262D" />
                    <XAxis
                      dataKey="month"
                      tick={{ fill: "#8B949E", fontSize: 11 }}
                      axisLine={false}
                      tickLine={false}
                    />
                    <YAxis
                      tick={{ fill: "#8B949E", fontSize: 11 }}
                      axisLine={false}
                      tickLine={false}
                    />
                    <Tooltip
                      contentStyle={{
                        background: "#1C2128",
                        border: "1px solid #30363D",
                        borderRadius: 8,
                        fontSize: 12,
                      }}
                    />
                    <Area
                      type="monotone"
                      dataKey="total"
                      stroke="#74B9FF"
                      strokeWidth={2}
                      fill="url(#g1)"
                      dot={{ fill: "#74B9FF", r: 3 }}
                      name="직원수"
                    />
                    <Bar
                      dataKey="join"
                      fill="#00D4AA"
                      radius={[3, 3, 0, 0]}
                      name="입사"
                    />
                    <Bar
                      dataKey="leave"
                      fill="#FF6B6B"
                      radius={[3, 3, 0, 0]}
                      name="퇴사"
                    />
                  </AreaChart>
                </ResponsiveContainer>
              </div>
              <div
                style={{
                  background: "#161B22",
                  border: "1px solid #21262D",
                  borderRadius: 12,
                  padding: "22px 24px",
                }}
              >
                <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>
                  부서별 인원
                </div>
                <div
                  style={{ color: "#8B949E", fontSize: 12, marginBottom: 8 }}
                >
                  전체 {totalEmp}명
                </div>
                <ResponsiveContainer width="100%" height={140}>
                  <PieChart>
                    <Pie
                      data={deptData}
                      cx="50%"
                      cy="50%"
                      innerRadius={40}
                      outerRadius={65}
                      dataKey="value"
                      paddingAngle={3}
                    >
                      {deptData.map((e, i) => (
                        <Cell key={i} fill={e.color} />
                      ))}
                    </Pie>
                    <Tooltip
                      contentStyle={{
                        background: "#1C2128",
                        border: "1px solid #30363D",
                        borderRadius: 8,
                        fontSize: 12,
                      }}
                    />
                  </PieChart>
                </ResponsiveContainer>
                <div
                  style={{
                    display: "grid",
                    gridTemplateColumns: "1fr 1fr",
                    gap: "4px 12px",
                  }}
                >
                  {deptData.map((d, i) => (
                    <div
                      key={i}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                        fontSize: 11,
                      }}
                    >
                      <div
                        style={{
                          width: 8,
                          height: 8,
                          borderRadius: 2,
                          background: d.color,
                          flexShrink: 0,
                        }}
                      />
                      <span style={{ color: "#8B949E" }}>{d.name}</span>
                      <span style={{ marginLeft: "auto", fontWeight: 700 }}>
                        {d.value}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </>
        )}

        {/* ── HEADCOUNT ──────────────────────────────── */}
        {activeTab === "headcount" && (
          <div
            style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}
          >
            <div
              style={{
                background: "#161B22",
                border: "1px solid #21262D",
                borderRadius: 12,
                padding: "24px",
              }}
            >
              <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>
                월별 입퇴사 현황
              </div>
              <div style={{ color: "#8B949E", fontSize: 12, marginBottom: 20 }}>
                추이
              </div>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={headcount}>
                  <CartesianGrid
                    strokeDasharray="3 3"
                    stroke="#21262D"
                    vertical={false}
                  />
                  <XAxis
                    dataKey="month"
                    tick={{ fill: "#8B949E", fontSize: 11 }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <YAxis
                    tick={{ fill: "#8B949E", fontSize: 11 }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <Tooltip
                    contentStyle={{
                      background: "#1C2128",
                      border: "1px solid #30363D",
                      borderRadius: 8,
                      fontSize: 12,
                    }}
                  />
                  <Bar
                    dataKey="join"
                    fill="#00D4AA"
                    radius={[4, 4, 0, 0]}
                    name="입사"
                  />
                  <Bar
                    dataKey="leave"
                    fill="#FF6B6B"
                    radius={[4, 4, 0, 0]}
                    name="퇴사"
                  />
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div
              style={{
                background: "#161B22",
                border: "1px solid #21262D",
                borderRadius: 12,
                padding: "24px",
              }}
            >
              <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>
                부서별 인원 분포
              </div>
              <div style={{ color: "#8B949E", fontSize: 12, marginBottom: 20 }}>
                전체 {totalEmp}명
              </div>
              <div
                style={{ display: "flex", flexDirection: "column", gap: 12 }}
              >
                {deptData.map((d, i) => (
                  <div key={i}>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                        marginBottom: 5,
                      }}
                    >
                      <span style={{ fontSize: 13, fontWeight: 600 }}>
                        {d.name}
                      </span>
                      <span
                        style={{
                          fontSize: 13,
                          fontWeight: 700,
                          color: d.color,
                        }}
                      >
                        {d.value}명
                      </span>
                    </div>
                    <div
                      style={{
                        background: "#21262D",
                        borderRadius: 4,
                        height: 8,
                      }}
                    >
                      <div
                        style={{
                          width: `${(d.value / totalEmp) * 100}%`,
                          height: "100%",
                          background: d.color,
                          borderRadius: 4,
                        }}
                      />
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ── PIPELINE ───────────────────────────────── */}
        {activeTab === "pipeline" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(4,1fr)",
                gap: 16,
              }}
            >
              {[
                {
                  label: "오픈 포지션",
                  value: `${totalOpenings}개`,
                  color: "#74B9FF",
                  icon: "📋",
                },
                {
                  label: "긴급 채용",
                  value: `${urgentCount}건`,
                  color: "#FF6B6B",
                  icon: "🚨",
                },
                {
                  label: "이번달 목표",
                  value: "5명",
                  color: "#FFD166",
                  icon: "🎯",
                },
                {
                  label: "평균 채용 기간",
                  value: "32일",
                  color: "#00D4AA",
                  icon: "⏳",
                },
              ].map((s, i) => (
                <div
                  key={i}
                  style={{
                    background: "#161B22",
                    border: "1px solid #21262D",
                    borderRadius: 12,
                    padding: "18px 20px",
                  }}
                >
                  <div style={{ fontSize: 20, marginBottom: 8 }}>{s.icon}</div>
                  <div
                    style={{ fontSize: 22, fontWeight: 800, color: s.color }}
                  >
                    {s.value}
                  </div>
                  <div style={{ fontSize: 12, color: "#8B949E", marginTop: 4 }}>
                    {s.label}
                  </div>
                </div>
              ))}
            </div>
            <div
              style={{
                background: "#161B22",
                border: "1px solid #21262D",
                borderRadius: 12,
                padding: "24px",
              }}
            >
              <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>
                포지션별 채용 파이프라인
              </div>
              <div style={{ color: "#8B949E", fontSize: 12, marginBottom: 20 }}>
                단계별 후보자 진행 현황
              </div>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "2fr 1fr 1fr 1fr 1fr 1fr 100px",
                  gap: 8,
                  padding: "8px 12px",
                  background: "#0D1117",
                  borderRadius: 8,
                  marginBottom: 10,
                }}
              >
                {[
                  "포지션",
                  "서류",
                  "1차면접",
                  "2차면접",
                  "최종",
                  "상태",
                  "진행률",
                ].map((h, i) => (
                  <div
                    key={i}
                    style={{
                      fontSize: 11,
                      fontWeight: 700,
                      color: "#8B949E",
                      textAlign: i > 0 ? "center" : "left",
                    }}
                  >
                    {h}
                  </div>
                ))}
              </div>
              {pipeline.map((row, i) => {
                const stages = [row.s1, row.s2, row.s3, row.s4];
                const maxStage =
                  row.s4 > 0 ? 4 : row.s3 > 0 ? 3 : row.s2 > 0 ? 2 : 1;
                const progress = Math.round((maxStage / 4) * 100);
                return (
                  <div
                    key={i}
                    style={{
                      display: "grid",
                      gridTemplateColumns: "2fr 1fr 1fr 1fr 1fr 1fr 100px",
                      gap: 8,
                      padding: "14px 12px",
                      borderBottom:
                        i < pipeline.length - 1 ? "1px solid #21262D" : "none",
                      alignItems: "center",
                      background: row.urgent ? "#FF6B6B08" : "transparent",
                      borderRadius: 6,
                    }}
                  >
                    <div>
                      <div
                        style={{
                          fontSize: 13,
                          fontWeight: 600,
                          display: "flex",
                          alignItems: "center",
                          gap: 6,
                        }}
                      >
                        {row.urgent && (
                          <span
                            style={{
                              fontSize: 9,
                              background: "#FF6B6B22",
                              color: "#FF6B6B",
                              borderRadius: 4,
                              padding: "1px 5px",
                              fontWeight: 700,
                            }}
                          >
                            긴급
                          </span>
                        )}
                        {row.position}
                      </div>
                      <div
                        style={{ fontSize: 10, color: "#8B949E", marginTop: 2 }}
                      >
                        총 {row.total}명 채용 목표
                      </div>
                    </div>
                    {stages.map((val, si) => (
                      <div key={si} style={{ textAlign: "center" }}>
                        <div
                          style={{
                            display: "inline-flex",
                            alignItems: "center",
                            justifyContent: "center",
                            width: 28,
                            height: 28,
                            borderRadius: "50%",
                            background:
                              val > 0 ? stageColors[si] + "22" : "#21262D",
                            color: val > 0 ? stageColors[si] : "#6E7681",
                            fontWeight: 700,
                            fontSize: 13,
                          }}
                        >
                          {val}
                        </div>
                      </div>
                    ))}
                    <div style={{ textAlign: "center" }}>
                      <span
                        style={{
                          fontSize: 10,
                          fontWeight: 700,
                          borderRadius: 6,
                          padding: "3px 8px",
                          background:
                            (statusColors[row.status] || "#8B949E") + "22",
                          color: statusColors[row.status] || "#8B949E",
                        }}
                      >
                        {row.status}
                      </span>
                    </div>
                    <div>
                      <div
                        style={{
                          background: "#21262D",
                          borderRadius: 4,
                          height: 6,
                          overflow: "hidden",
                        }}
                      >
                        <div
                          style={{
                            width: `${progress}%`,
                            height: "100%",
                            background:
                              progress === 100
                                ? "#00D4AA"
                                : progress >= 75
                                ? "#FFD166"
                                : "#74B9FF",
                            borderRadius: 4,
                          }}
                        />
                      </div>
                      <div
                        style={{
                          fontSize: 10,
                          color: "#8B949E",
                          marginTop: 4,
                          textAlign: "right",
                        }}
                      >
                        {progress}%
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ── BUDGET ─────────────────────────────────── */}
        {activeTab === "budget" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(4,1fr)",
                gap: 16,
              }}
            >
              {[
                {
                  label: "최근월 예산",
                  value: `${lastBudget.budget.toLocaleString()}만원`,
                  color: "#74B9FF",
                  icon: "📊",
                  sub: "",
                },
                {
                  label: "최근월 실적",
                  value: `${lastBudget.actual.toLocaleString()}만원`,
                  color: Number(budgetRate) > 100 ? "#FF6B6B" : "#00D4AA",
                  icon: "💰",
                  sub: Number(budgetRate) > 100 ? "예산 초과" : "예산 내",
                },
                {
                  label: "예산 집행률",
                  value: `${budgetRate}%`,
                  color: Number(budgetRate) > 100 ? "#FF6B6B" : "#00D4AA",
                  icon: "📈",
                  sub: Number(budgetRate) > 100 ? "⚠️ 초과" : "✅ 정상",
                },
                {
                  label: "항목 수",
                  value: `${breakdown.length}개`,
                  color: "#A29BFE",
                  icon: "🏦",
                  sub: "인건비 항목",
                },
              ].map((s, i) => (
                <div
                  key={i}
                  style={{
                    background: "#161B22",
                    border: "1px solid #21262D",
                    borderRadius: 12,
                    padding: "18px 20px",
                    position: "relative",
                    overflow: "hidden",
                  }}
                >
                  <div
                    style={{
                      position: "absolute",
                      top: 0,
                      right: 0,
                      width: 70,
                      height: 70,
                      borderRadius: "0 12px 0 70px",
                      background: s.color + "0D",
                    }}
                  />
                  <div style={{ fontSize: 20, marginBottom: 8 }}>{s.icon}</div>
                  <div
                    style={{ fontSize: 20, fontWeight: 800, color: s.color }}
                  >
                    {s.value}
                  </div>
                  <div style={{ fontSize: 12, color: "#8B949E", marginTop: 4 }}>
                    {s.label}
                  </div>
                  {s.sub && (
                    <div
                      style={{
                        fontSize: 11,
                        color: "#6E7681",
                        marginTop: 6,
                        borderTop: "1px solid #21262D",
                        paddingTop: 6,
                      }}
                    >
                      {s.sub}
                    </div>
                  )}
                </div>
              ))}
            </div>
            <div
              style={{
                background: "#161B22",
                border: "1px solid #21262D",
                borderRadius: 12,
                padding: "24px",
              }}
            >
              <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>
                월별 인건비 예산 vs 실적
              </div>
              <div style={{ color: "#8B949E", fontSize: 12, marginBottom: 20 }}>
                단위: 만원
              </div>
              <ResponsiveContainer width="100%" height={220}>
                <LineChart data={budget}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#21262D" />
                  <XAxis
                    dataKey="month"
                    tick={{ fill: "#8B949E", fontSize: 11 }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <YAxis
                    tick={{ fill: "#8B949E", fontSize: 11 }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <Tooltip
                    contentStyle={{
                      background: "#1C2128",
                      border: "1px solid #30363D",
                      borderRadius: 8,
                      fontSize: 12,
                    }}
                    formatter={(v: number) => `${v.toLocaleString()}만원`}
                  />
                  <Line
                    type="monotone"
                    dataKey="budget"
                    stroke="#74B9FF"
                    strokeWidth={2}
                    strokeDasharray="6 3"
                    dot={{ fill: "#74B9FF", r: 4 }}
                    name="예산"
                  />
                  <Line
                    type="monotone"
                    dataKey="actual"
                    stroke="#FF6B6B"
                    strokeWidth={2}
                    dot={{ fill: "#FF6B6B", r: 4 }}
                    name="실적"
                  />
                </LineChart>
              </ResponsiveContainer>
            </div>
            <div
              style={{
                background: "#161B22",
                border: "1px solid #21262D",
                borderRadius: 12,
                padding: "24px",
              }}
            >
              <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>
                항목별 인건비 현황
              </div>
              <div style={{ color: "#8B949E", fontSize: 12, marginBottom: 20 }}>
                단위: 만원
              </div>
              <div
                style={{ display: "flex", flexDirection: "column", gap: 14 }}
              >
                {breakdown.map((b, i) => {
                  const rate = ((b.actual / b.budget) * 100).toFixed(1);
                  const over = b.actual > b.budget;
                  return (
                    <div key={i}>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          marginBottom: 6,
                        }}
                      >
                        <div
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: 8,
                          }}
                        >
                          <div
                            style={{
                              width: 10,
                              height: 10,
                              borderRadius: 2,
                              background: b.color,
                            }}
                          />
                          <span style={{ fontSize: 13, fontWeight: 600 }}>
                            {b.label}
                          </span>
                        </div>
                        <div style={{ display: "flex", gap: 16, fontSize: 12 }}>
                          <span style={{ color: "#8B949E" }}>
                            예산 {b.budget.toLocaleString()}
                          </span>
                          <span
                            style={{
                              color: over ? "#FF6B6B" : "#00D4AA",
                              fontWeight: 700,
                            }}
                          >
                            실적 {b.actual.toLocaleString()}
                          </span>
                          <span
                            style={{
                              color: over ? "#FF6B6B" : "#00D4AA",
                              fontWeight: 700,
                            }}
                          >
                            {rate}%
                          </span>
                        </div>
                      </div>
                      <div
                        style={{
                          background: "#21262D",
                          borderRadius: 4,
                          height: 8,
                          overflow: "hidden",
                        }}
                      >
                        <div
                          style={{
                            width: `${Math.min(Number(rate), 110)}%`,
                            height: "100%",
                            background: over ? "#FF6B6B" : b.color,
                            borderRadius: 4,
                          }}
                        />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
