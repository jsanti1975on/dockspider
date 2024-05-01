#include <stdio.h>   // For printf()
#include <stdlib.h>  // For system()

int main() {
    const char* command = "netsh interface ipv4 add dns Ethernet 8.8.8.8 index=2";

    // Execute the command using system()
    int result = system(command);

    // Check the result of the system call
    if (result == 0) {
        printf("Command executed successfully.\n");
    } else {
        printf("Error executing command.\n");
    }

    return 0;
}
