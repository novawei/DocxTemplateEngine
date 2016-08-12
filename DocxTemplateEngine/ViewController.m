//
//  ViewController.m
//  DocxTemplateEngine
//
//  Created by WeiXinxing on 16/8/12.
//  Copyright © 2016年 nfs. All rights reserved.
//

#import "ViewController.h"
#import "DocxTemplateEngine.h"

@interface ViewController () <UIDocumentInteractionControllerDelegate>

@property (nonatomic, strong) UIDocumentInteractionController *documentController;

@end

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
    // Do any additional setup after loading the view, typically from a nib.
    dispatch_async(dispatch_get_global_queue(DISPATCH_QUEUE_PRIORITY_DEFAULT, 0), ^{
        NSString *path = [self createDocx];
        NSLog(@"%@", path);
        dispatch_async(dispatch_get_main_queue(), ^{
            [self showInOtherApp:path];
        });
    });
}

- (void)didReceiveMemoryWarning {
    [super didReceiveMemoryWarning];
    // Dispose of any resources that can be recreated.
}

- (NSString *)createDocx
{
    NSString *path = [DocxTemplateEngine createDocxWithName:@"test"
                                                       data:@{@"title": @"测试标题",
                                                              @"items": @[@"条目1", @"条目2", @"条目3"]
                                                              }
                                              usingTemplate:@"template2"];
    return path;
//    NSString *path = [DocxTemplateEngine createDocxWithName:@"请假单"
//                                                       data:@{@"WRITE_DATE": @"2016-08-12",
//                                                              @"NAME": @"魏新星",
//                                                              @"DEPARTMENT": @"宇宙研发中心",
//                                                              @"POSITION": @"闲人",
//                                                              @"START_Y": @"2016",
//                                                              @"START_M": @"08",
//                                                              @"START_D": @"15",
//                                                              @"START_H": @"08",
//                                                              @"END_Y": @"2016",
//                                                              @"END_M": @"08",
//                                                              @"END_D": @"19",
//                                                              @"END_H": @"18",
//                                                              @"COUNT": @"5",
//                                                              @"REASON": @"家里没酱油了，打酱油去"
//                                                              }
//                                              usingTemplate:@"template1"];
//    return path;
}

- (void)showInOtherApp:(NSString *)path
{
    if (!path) {
        return;
    }
    
    self.documentController = [UIDocumentInteractionController interactionControllerWithURL:[NSURL fileURLWithPath:path]];
    self.documentController.delegate = self;
    [self.documentController presentOpenInMenuFromRect:self.view.bounds inView:self.view animated:YES];
}

#pragma mark - UIDocumentInteractionControllerDelegate

- (void)documentInteractionControllerDidDismissOpenInMenu:(UIDocumentInteractionController *)controller
{
    self.documentController = nil;
}


@end
