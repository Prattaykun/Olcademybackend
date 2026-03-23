import * as fs from "fs";
import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, BorderStyle } from "docx";

const apiData = [
    {
        module: "User APIs",
        prefix: "/user",
        endpoints: [
            { method: "POST", path: "/signup", description: "Register a new user" },
            { method: "POST", path: "/login", description: "User login" },
            { method: "POST", path: "/signup/verify-email", description: "Verify email with OTP" },
            { method: "POST", path: "/signup/resend-otp", description: "Resend verification OTP" },
            { method: "GET", path: "/check-username/:username", description: "Check if username is available" },
            { method: "GET", path: "/profile", description: "Get user profile" },
            { method: "PUT", path: "/profile", description: "Update user profile" },
            { method: "GET", path: "/shipment", description: "Get shipment info" },
            { method: "PUT", path: "/shipment", description: "Update shipment info" },
            { method: "GET", path: "/admin/users", description: "Get all users (Admin)" },
            { method: "PATCH", path: "/admin/users/:userId/admin-status", description: "Update user admin status (Admin)" }
        ]
    },
    {
        module: "Order APIs",
        prefix: "/order",
        endpoints: [
            { method: "POST", path: "/create", description: "Create a new order" },
            { method: "GET", path: "/:orderNumber", description: "Get order details" },
            { method: "GET", path: "/:orderNumber/invoice", description: "Get order invoice" },
            { method: "GET", path: "/email/:email", description: "Get orders by email" },
            { method: "GET", path: "/user/my-orders", description: "Get current user's orders" },
            { method: "GET", path: "/", description: "Get all orders (Admin)" },
            { method: "GET", path: "/admin/statistics", description: "Get order statistics (Admin)" },
            { method: "PUT", path: "/:orderNumber/status", description: "Update order status (Admin)" },
            { method: "PATCH", path: "/admin/bulk-update-status", description: "Bulk update order status (Admin)" },
            { method: "DELETE", path: "/:orderNumber", description: "Delete order (Admin)" },
            { method: "PATCH", path: "/:orderNumber/tracking", description: "Update tracking info (Admin)" },
            { method: "PATCH", path: "/:orderNumber/admin-notes", description: "Update admin notes (Admin)" },
            { method: "GET", path: "/admin/status/:status", description: "Get orders by status (Admin)" },
            { method: "GET", path: "/admin/export/csv", description: "Export orders to CSV (Admin)" },
            { method: "GET", path: "/admin/analytics", description: "Get order analytics (Admin)" },
            { method: "GET", path: "/admin/search", description: "Search orders (Admin)" },
            { method: "POST", path: "/:orderNumber/review", description: "Add a review for an order" },
            { method: "GET", path: "/admin/reviews", description: "Get all reviews (Admin)" }
        ]
    },
    {
        module: "Cart APIs",
        prefix: "/cart",
        endpoints: [
            { method: "GET", path: "/", description: "Get current user's cart" },
            { method: "POST", path: "/add", description: "Add item to cart" },
            { method: "DELETE", path: "/remove/:id", description: "Remove item from cart" },
            { method: "PUT", path: "/update/:id", description: "Update cart item quantity" },
            { method: "DELETE", path: "/clear", description: "Clear cart" }
        ]
    },
    {
        module: "Wishlist APIs",
        prefix: "/wishlist",
        endpoints: [
            { method: "GET", path: "/", description: "Get current user's wishlist" },
            { method: "POST", path: "/add", description: "Add item to wishlist" },
            { method: "DELETE", path: "/remove/:id", description: "Remove item from wishlist" },
            { method: "POST", path: "/move-to-cart/:id", description: "Move item from wishlist to cart" },
            { method: "DELETE", path: "/clear", description: "Clear wishlist" }
        ]
    },
    {
        module: "Product APIs",
        prefix: "/api/products",
        endpoints: [
            { method: "GET", path: "/:id/related", description: "Get related products" },
            { method: "GET", path: "/search", description: "Search products" },
            { method: "GET", path: "/gifts/collections", description: "Get gifts collections" },
            { method: "GET", path: "/women/collections", description: "Get women collections" },
            { method: "GET", path: "/men/collections", description: "Get men collections" },
            { method: "GET", path: "/unisex/collections", description: "Get unisex collections" },
            { method: "GET", path: "/home/collections", description: "Get home collections" },
            { method: "GET", path: "/debug/count", description: "Get products count" },
            { method: "GET", path: "/debug/exists/:id", description: "Check if product exists" },
            { method: "GET", path: "/", description: "Get all products" },
            { method: "POST", path: "/", description: "Create a new product" },
            { method: "GET", path: "/:id", description: "Get product by ID" },
            { method: "PUT", path: "/:id", description: "Update product by ID" },
            { method: "DELETE", path: "/:id", description: "Delete product by ID" },
            { method: "POST", path: "/:id/toggle-status", description: "Toggle product status" },
            { method: "DELETE", path: "/:id/image/:index", description: "Delete product image" },
            { method: "DELETE", path: "/:id/hover-image", description: "Delete product hover image" }
        ]
    },
    {
        module: "Scent APIs",
        prefix: "/api/scents",
        endpoints: [
            { method: "GET", path: "/featured", description: "Get featured scents" },
            { method: "GET", path: "/search", description: "Search scents" },
            { method: "GET", path: "/collection/:collection", description: "Get scents by collection" },
            { method: "GET", path: "/brand/:brand", description: "Get scents by brand" },
            { method: "GET", path: "/filters", description: "Get scents filters" },
            { method: "GET", path: "/", description: "Get all scents" },
            { method: "POST", path: "/", description: "Create a new scent" },
            { method: "GET", path: "/:id/related", description: "Get related scents" },
            { method: "GET", path: "/:id", description: "Get scent by ID" },
            { method: "PUT", path: "/:id", description: "Update scent by ID" },
            { method: "DELETE", path: "/:id", description: "Delete scent by ID" },
            { method: "PATCH", path: "/:id/deactivate", description: "Deactivate scent by ID" },
            { method: "DELETE", path: "/:id/images/:index", description: "Delete scent image" },
            { method: "DELETE", path: "/:id/hover-image", description: "Delete scent hover image" }
        ]
    },
    {
        module: "Banner APIs",
        prefix: "/api/banners",
        endpoints: [
            { method: "GET", path: "/debug/count", description: "Get banners count" },
            { method: "GET", path: "/", description: "Get all banners" },
            { method: "POST", path: "/", description: "Create a new banner" },
            { method: "PUT", path: "/:id", description: "Update banner by ID" },
            { method: "DELETE", path: "/:id", description: "Delete banner by ID" },
            { method: "POST", path: "/:id/toggle-status", description: "Toggle banner status" },
            { method: "DELETE", path: "/:id/image/:imageType", description: "Delete banner image" },
            { method: "POST", path: "/:id/click", description: "Record banner click" },
            { method: "GET", path: "/:category", description: "Get banners by category" },
            { method: "GET", path: "/:category/:type", description: "Get banners by category and type" }
        ]
    },
    {
        module: "Subscriber APIs",
        prefix: "/api",
        endpoints: [
            { method: "POST", path: "/subscribe", description: "Subscribe to newsletter" }
        ]
    },
    {
        module: "Test and Debug APIs",
        prefix: "",
        endpoints: [
            { method: "GET", path: "/test-scents", description: "Test scents endpoint" },
            { method: "GET", path: "/test-scent-creation", description: "Test scent creation endpoint" },
            { method: "GET", path: "/test-trending", description: "Test trending scents endpoint" },
            { method: "GET", path: "/test-bestsellers", description: "Test best sellers endpoint" },
            { method: "GET", path: "/test-womens-signature", description: "Test women's signature endpoint" },
            { method: "GET", path: "/test-rose-garden-essence", description: "Test rose garden essence endpoint" },
            { method: "GET", path: "/test-mens-signature", description: "Test men's signature endpoint" },
            { method: "GET", path: "/test-orange-marmalade", description: "Test orange marmalade endpoint" },
            { method: "GET", path: "/test-gender-free", description: "Test gender-free endpoint" },
            { method: "GET", path: "/test-limitless", description: "Test limitless endpoint" },
            { method: "GET", path: "/test-perfect-discover-gifts", description: "Test perfect discover gifts endpoint" },
            { method: "GET", path: "/test-perfect-gifts-premium", description: "Test perfect gifts premium endpoint" },
            { method: "GET", path: "/test-perfect-gifts-luxury", description: "Test perfect gifts luxury endpoint" },
            { method: "GET", path: "/test-home-decor-gifts", description: "Test home decor gifts endpoint" },
            { method: "GET", path: "/test-wishlist", description: "Test wishlist endpoint" },
            { method: "GET", path: "/test-orders", description: "Test orders endpoint" },
            { method: "GET", path: "/test-products", description: "Test products endpoint" },
            { method: "GET", path: "/test-banners", description: "Test banners endpoint" },
            { method: "GET", path: "/debug/static", description: "Debug static files endpoint" },
            { method: "GET", path: "/", description: "Health check / Root endpoint" },
            { method: "GET", path: "/debug/routes", description: "Debug routes endpoint" }
        ]
    }
];

function createApiTable(endpoints, prefix) {
    const tableRows = [
        new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Method", bold: true })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Full Path", bold: true })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Description", bold: true })] })] })
            ],
            tableHeader: true
        })
    ];

    endpoints.forEach(ep => {
        const fullPath = prefix + ep.path;
        tableRows.push(
            new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ text: ep.method })] }),
                    new TableCell({ children: [new Paragraph({ text: fullPath })] }),
                    new TableCell({ children: [new Paragraph({ text: ep.description || "" })] })
                ]
            })
        );
    });

    return new Table({
        rows: tableRows,
        width: { size: 100, type: "pct" },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 1 },
            bottom: { style: BorderStyle.SINGLE, size: 1 },
            left: { style: BorderStyle.SINGLE, size: 1 },
            right: { style: BorderStyle.SINGLE, size: 1 },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
            insideVertical: { style: BorderStyle.SINGLE, size: 1 },
        }
    });
}

const docChildren = [
    new Paragraph({
        text: "OLCAcademy API Documentation",
        heading: HeadingLevel.TITLE,
    }),
    new Paragraph({
        text: "This document contains a comprehensive list of all the APIs available in the backend system along with their methods and paths for testing.",
        spacing: { after: 400 }
    })
];

apiData.forEach(moduleData => {
    docChildren.push(
        new Paragraph({
            text: moduleData.module,
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 400, after: 200 }
        })
    );

    if (moduleData.prefix) {
        docChildren.push(
            new Paragraph({
                text: "Route Prefix: " + moduleData.prefix,
                spacing: { after: 200 }
            })
        );
    }

    docChildren.push(createApiTable(moduleData.endpoints, moduleData.prefix));
});

const doc = new Document({
    sections: [{
        properties: {},
        children: docChildren
    }]
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("api_documentation.docx", buffer);
    console.log("Document created successfully at api_documentation.docx");
}).catch(err => {
    console.error("Error creating document", err);
});
