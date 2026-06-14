export interface VbaSample {
  id: string;
  title: string;
  description: string;
  category: 'excel' | 'outlook' | 'file' | 'utility' | 'speed';
  code: string;
  explanation: string[];
  difficulty: 'easy' | 'medium' | 'hard';
  tags: string[];
}
