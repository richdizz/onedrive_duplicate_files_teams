import File from "./file";

export interface Duplicate {
    fileName: string;
    fileExt: string;
    size: number;
    locations: File[];
    fileToKeep: string;
}