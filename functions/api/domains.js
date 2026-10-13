export async function onRequestGet(context) {
    // Database se master domains fetch karega
    try {
        const { results } = await context.env.DB.prepare("SELECT domain FROM master_sites").all();
        const domains = results.map(r => r.domain);
        return new Response(JSON.stringify(domains), { 
            headers: { "Content-Type": "application/json" } 
        });
    } catch (e) {
        return new Response(e.message, { status: 500 });
    }
}

export async function onRequestPost(context) {
    // Naye unique domains ko database mein save karega
    try {
        const body = await context.request.json();
        const domains = body.domains; 
        
        if (!domains || !domains.length) {
            return new Response("No domains provided", { status: 400 });
        }

        // Batch insert to handle multiple domains fast
        const statements = domains.map(domain =>
            context.env.DB.prepare("INSERT OR IGNORE INTO master_sites (domain) VALUES (?)").bind(domain)
        );
        
        await context.env.DB.batch(statements);

        return new Response(JSON.stringify({ success: true, added: domains.length }), { 
            headers: { "Content-Type": "application/json" } 
        });
    } catch (e) {
        return new Response(e.message, { status: 500 });
    }
}