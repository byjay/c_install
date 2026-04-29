// CF Pages Function: /api/config
// 환경변수 → 클라이언트에 공개 (ADMIN_PASSWORD 제외)
export async function onRequest(context) {
  const { env } = context;
  return Response.json(
    {
      users: [env.USER1, env.USER2].filter(Boolean),
      userPassword: env.USER_PASSWORD || "",
      adminPassword: env.ADMIN_PASSWORD || "",
    },
    { headers: { "Cache-Control": "no-store, no-cache" } }
  );
}
