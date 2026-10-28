import DataChecker from "./components/DataChecker";

export default function Home() {
  return (
    <div className="min-h-screen flex flex-col items-center px-6 py-16 gap-10">
      <div className="text-center">
        <h1 className="text-3xl font-bold tracking-tight">Loan Data Checker</h1>
        <p className="text-gray-500 mt-2">
          Upload a loan data file to validate it — file, cell, row and cross-row checks.
        </p>
      </div>
      <DataChecker />
    </div>
  );
}
