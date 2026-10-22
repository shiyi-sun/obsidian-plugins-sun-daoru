export default async function (ctx) {
  await ctx.reloadApp();
  await ctx.sleep(800);
  return { reloaded: true };
}
