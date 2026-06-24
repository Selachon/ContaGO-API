import "dotenv/config";
import { connectMongo } from "./src/services/database.js";
import { listCompanies, getCompanyContext } from "./src/services/siigoCompaniesService.js";
import { runWithSiigoCompany, listAccounts } from "./src/services/siigoService.js";
async function main(){
  await connectMongo();
  const comps = await listCompanies();
  const ctx = await getCompanyContext(comps[0].id);
  await runWithSiigoCompany(ctx!, async () => {
    const accs:any = await listAccounts();
    const arr = Array.isArray(accs) ? accs : (accs?.results || accs?.data || []);
    console.log("tipo respuesta:", Array.isArray(accs)?"array":"obj", "| total cuentas:", arr.length);
    if(arr[0]) console.log("\nclaves cuenta:", Object.keys(arr[0]));
    console.log("\n=== 5 cuentas crudas ===");
    console.log(JSON.stringify(arr.slice(0,5), null, 2));
  });
}
main().then(()=>process.exit(0)).catch(e=>{console.error("ERR", e?.status, e?.message);process.exit(1);});
