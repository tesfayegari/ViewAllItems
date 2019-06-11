
export interface IPagingProps {
    totalItems: number;
    itemsCountPerPage: number;
    onPageUpdate: (pageNumber: number) => void;
    currentPage: number;
}