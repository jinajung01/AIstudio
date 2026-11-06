export interface Application {
  id: string;
  code: string;
  name: string;
  phone: string;
  email: string;
  slot: string;
  memo?: string;
  agree: boolean;
  createdAt: string;
  updatedAt: string;
}

export type ApplicationInput = Omit<Application, 'id' | 'createdAt' | 'updatedAt' | 'code'> & {
  id?: string;
  code?: string;
};
