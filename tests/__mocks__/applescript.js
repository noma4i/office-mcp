export const mockExecFile = {
  stdout: '',
  stderr: '',
};

export function setMockResponse(stdout, stderr = '') {
  mockExecFile.stdout = stdout;
  mockExecFile.stderr = stderr;
}

export function createMockExecFile() {
  return jest.fn((command, args) => {
    if (mockExecFile.stderr) {
      return Promise.reject(new Error(mockExecFile.stderr));
    }
    return Promise.resolve({ stdout: mockExecFile.stdout, stderr: '' });
  });
}

export function resetMock() {
  mockExecFile.stdout = '';
  mockExecFile.stderr = '';
}
