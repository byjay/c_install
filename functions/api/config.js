// CF Pages Function: /api/config
// 환경변수 → 클라이언트에 공개 (비밀값 포함: 내부 전용 빌드)
export async function onRequest(context) {
  const { env } = context;
  return Response.json(
    {
      users: [env.USER1, env.USER2].filter(Boolean),
      userPassword: env.USER_PASSWORD || "",
      adminId: env.ADMIN_id || "",         // 관리자 ID (ADMIN_id)
      adminPassword: env.ADMIN_PASSWORD || "",
    },
    { headers: { "Cache-Control": "no-store, no-cache" } }
  );
}
