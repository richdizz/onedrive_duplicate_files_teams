import { Duplicate } from "./duplicate";

export interface Scan {
    id: string;
    user: string;
    scanDate: string;
    status: string;
    duplicates: Duplicate[];
}