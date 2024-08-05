// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Importing the @azure/search-documents library
import {
    AzureKeyCredential,
    ComplexField,
    odata,
    SearchClient,
    SearchFieldArray,
    SearchIndex,
    SearchIndexClient,
    SearchSuggester,
    SimpleField
} from "@azure/search-documents";

// Load the .env file if it exists
import dotenv from 'dotenv';
dotenv.config();

// Importing the index definition and sample data
import hotelData from './hotels.json';
import indexDefinition from './hotels_quickstart_index.json';
if (!hotelData || !indexDefinition) {
    throw new Error("Make sure to set valid values for hotelData and indexDefinition.");
}

interface Hotel {
    HotelId: string;
    HotelName: string;
    Description: string;
    Description_fr: string;
    Category: string;
    Tags: string[];
    ParkingIncluded: string | boolean;
    LastRenovationDate: string;
    Rating: number;
    Address: {
        StreetAddress: string;
        City: string;
        StateProvince: string;
        PostalCode: string;
    };
};

interface HotelIndexDefinition {
    name: string;
    fields: SimpleField[] | ComplexField[];
    suggesters: SearchSuggester[];
};
const hotels: Hotel[] = hotelData["value"];
const hotelIndexDefinition: HotelIndexDefinition = indexDefinition as HotelIndexDefinition;

// Getting endpoint and apiKey from .env file
const endpoint = process.env.SEARCH_API_ENDPOINT || "";
const apiKey = process.env.SEARCH_API_KEY || "";

async function main(): Promise<void> {
    console.log(`Running Azure AI Search Javascript quickstart...`);
    if (!endpoint || !apiKey) {
        console.log("Make sure to set valid values for endpoint and apiKey with proper authorization.");
        return;
    }

    // Creating an index client to create the search index
    const indexClient = new SearchIndexClient(endpoint, new AzureKeyCredential(apiKey));

    // Getting the name of the index from the index definition
    const indexName: string = hotelIndexDefinition.name;

    console.log('Checking if index exists...');
    await deleteIndexIfExists(indexClient, indexName);

    console.log('Creating index...');
    let index = await indexClient.createIndex(hotelIndexDefinition);
    console.log(`Index named ${index.name} has been created.`);

    // Creating a search client to upload documents and issue queries
    const searchClient = indexClient.getSearchClient<Hotel>(indexName);

    // Load the data
    await loadData(searchClient, hotels);

    // waiting one second for indexing to complete (for demo purposes only)
    await sleep(1000);

    // Verify docs are indexed
    const documentCount = await searchClient.getDocumentsCount();
    console.log(
        `${documentCount} docs uploaded`,
    );

    console.log('Querying the index...');
    console.log();

    await sendQueries(searchClient);
}

async function loadData(
    searchClient: SearchClient<Hotel>,
    hotels: Hotel[],
): Promise<void> {
    console.log("Uploading documents...");

    const indexDocumentsResult = await searchClient.mergeOrUploadDocuments(hotels);

    console.log(JSON.stringify(indexDocumentsResult));

    console.log(
        `Index operations succeeded: ${JSON.stringify(indexDocumentsResult.results[0].succeeded)}`,
    );
}

async function deleteIndexIfExists(indexClient: SearchIndexClient, indexName: string): Promise<void> {
    try {
        await indexClient.deleteIndex(indexName);
        console.log('Deleting index...');
    } catch {
        console.log('Index does not exist yet.');
    }
}

async function sendQueries(searchClient: SearchClient<Hotel>): Promise<void> {

    // Query 1
    console.log('Query #1 - search everything:');
    const selectFields: SearchFieldArray<Hotel> = [
        "HotelId",
        "HotelName",
        "Rating",
    ];
    const searchOptions1 = { 
        includeTotalCount: true, 
        select: selectFields 
    };

    let searchResults = await searchClient.search("*", searchOptions1);
    for await (const result of searchResults.results) {
        console.log(`${JSON.stringify(result.document)}`);
    }
    console.log(`Result count: ${searchResults.count}`);
    console.log();


    // Query 2
    console.log('Query #2 - search with filter, orderBy, and select:');
    let state = 'FL';
    const searchOptions2 = {
        filter: odata`Address/StateProvince eq ${state}`,
        orderBy: ["Rating desc"],
        select: selectFields
    };
    searchResults = await searchClient.search("wifi", searchOptions2);
    for await (const result of searchResults.results) {
        console.log(`${JSON.stringify(result.document)}`);
    }
    console.log();

    // Query 3
    console.log('Query #3 - limit searchFields:');
    const searchOptions3 = {
        select: selectFields,
        searchFields: ["HotelName"] as const
    };

    searchResults = await searchClient.search("sublime cliff", searchOptions3);
    for await (const result of searchResults.results) {
        console.log(`${JSON.stringify(result.document)}`);
    }
    console.log();

    // Query 4
    console.log('Query #4 - limit searchFields and use facets:');
    const searchOptions4 = {
        facets: ["Category"],
        select: selectFields,
        searchFields: ["HotelName"] as const
    };

    searchResults = await searchClient.search("*", searchOptions4);
    for await (const result of searchResults.results) {
        console.log(`${JSON.stringify(result.document)}`);
    }
    console.log();

    // Query 5
    console.log('Query #5 - Lookup document:');
    let documentResult = await searchClient.getDocument('3')
    console.log(`HotelId: ${documentResult.HotelId}; HotelName: ${documentResult.HotelName}`)
    console.log();
}

function sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
}
main().catch((err) => {
    console.error("The sample encountered an error:", err);
});