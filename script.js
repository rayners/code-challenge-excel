#!/usr/bin/env node
const args = require("yargs")
    .default("h", "127.0.0.1")
    .default("u", "root")
    .default("p", "password")
    .default("d", "code-challenge-excel")
    .default("o", "output.xlsx").argv;

const XLSX = require("xlsx");
const _ = require("lodash");
const knex = require("knex")({
    client: "mysql",
    connection: {
        host: args.h,
        user: args.u,
        password: args.p,
        database: args.d,
    },
});

const zones = {
    domestic: [1, 2, 3, 4, 5, 6, 7, 8],
    international: [
        "A",
        "B",
        "C",
        "D",
        "E",
        "F",
        "G",
        "H",
        "I",
        "J",
        "K",
        "L",
        "M",
        "N",
        "O",
    ],
};
const tabs = [
    {
        query: { shipping_speed: "standard", locale: "domestic" },
        title: "Domestic Standard Rates",
        zones: zones.domestic,
    },
    {
        query: { shipping_speed: "expedited", locale: "domestic" },
        title: "Domestic Expedited Rates",
        zones: zones.domestic,
    },
    {
        query: { shipping_speed: "nextDay", locale: "domestic" },
        title: "Domestic Next Day Rates",
        zones: zones.domestic,
    },
    {
        query: { shipping_speed: "intlEconomy", locale: "international" },
        title: "International Economy Rates",
        zones: zones.international,
    },
    {
        query: { shipping_speed: "intlExpedited", locale: "international" },
        title: "International Expedited Rates",
        zones: zones.international,
    },
];

const output = XLSX.utils.book_new();

(async () => {
    await Promise.all(
        tabs.map(async (tab) => {
            try {
                const res = await knex("rates")
                    .where({
                        client_id: 1240,
                        ...tab.query,
                    })
                    .select(
                        "start_weight",
                        "end_weight",
                        ...tab.zones.map((z) =>
                            knex.raw(
                                `sum(if(zone = "${z}", rate, 0)) as zone_${z}`
                            )
                        )
                    )
                    .groupBy("start_weight", "end_weight")
                    .orderBy("start_weight", "end_weight");

                const columns = [
                    "start_weight",
                    "end_weight",
                    ...tab.zones.map((z) => `zone_${z}`),
                ];
                const outputTab = XLSX.utils.json_to_sheet(
                    [
                        {
                            start_weight: "Start Weight",
                            end_weight: "End Weight",
                            ...tab.zones.reduce((acc, z) => {
                                acc[`zone_${z}`] = `Zone ${z}`;
                                return acc;
                            }, {}),
                        },
                        ...res,
                    ],
                    {
                        header: columns,
                        skipHeader: true,
                    }
                );

                // add some formatting (have to do it cell by cell :-/ )
                _.range(0, columns.length).forEach((col) => {
                    _.range(1, res.length + 1).forEach((row) => {
                        outputTab[
                            XLSX.utils.encode_cell({ c: col, r: row })
                        ].z = "0.00";
                    });
                });

                XLSX.utils.book_append_sheet(output, outputTab, tab.title);
            } catch (e) {
                console.log(e);
            }
        })
    );

    knex.destroy();
    XLSX.writeFile(output, args.o);
    return;
})();
