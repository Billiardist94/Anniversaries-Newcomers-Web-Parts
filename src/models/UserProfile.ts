export default interface UserProfile {
    name: string;
    firstName?: string;
    lastName?: string;
    phone?: string;
    mobile?: string;
    email: string;
    photoUrl: string;
    function: string;
    department?: string;
    skills?: string;
    projects?: string;
    office?: string;
    hireDate?: Date;
    internalHireDate?: Date;
}