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
    CGFloat x = (CGRectGetWidth(self.view.bounds)-200)*0.5;
    UIButton *button1 = [UIButton buttonWithType:UIButtonTypeRoundedRect];
    button1.frame = CGRectMake(x, 60, 200, 44);
    button1.titleLabel.font = [UIFont boldSystemFontOfSize:18];
    [button1 setTitle:@"模板一" forState:UIControlStateNormal];
    [button1 addTarget:self action:@selector(handleAction1:) forControlEvents:UIControlEventTouchUpInside];
    [self.view addSubview:button1];
    
    UIButton *button2 = [UIButton buttonWithType:UIButtonTypeRoundedRect];
    button2.frame = CGRectMake(x, 124, 200, 44);
    button2.titleLabel.font = [UIFont boldSystemFontOfSize:18];
    [button2 setTitle:@"模板二" forState:UIControlStateNormal];
    [button2 addTarget:self action:@selector(handleAction2:) forControlEvents:UIControlEventTouchUpInside];
    [self.view addSubview:button2];
}

- (void)handleAction1:(UIButton *)btn
{
    btn.enabled = NO;
    dispatch_async(dispatch_get_global_queue(DISPATCH_QUEUE_PRIORITY_DEFAULT, 0), ^{
        NSString *path = [self createDocx1];
        NSLog(@"%@", path);
        dispatch_async(dispatch_get_main_queue(), ^{
            [self showInOtherApp:path];
            btn.enabled = YES;
        });
    });
}

- (void)handleAction2:(UIButton *)btn
{
    btn.enabled = NO;
    dispatch_async(dispatch_get_global_queue(DISPATCH_QUEUE_PRIORITY_DEFAULT, 0), ^{
        NSString *path = [self createDocx2];
        NSLog(@"%@", path);
        dispatch_async(dispatch_get_main_queue(), ^{
            [self showInOtherApp:path];
            btn.enabled = YES;
        });
    });
}

- (void)didReceiveMemoryWarning {
    [super didReceiveMemoryWarning];
    // Dispose of any resources that can be recreated.
}

- (NSString *)createDocx1
{
    NSString *path = [DocxTemplateEngine createDocxWithName:@"请假单"
                                                       data:@{@"write_date": @"2016-08-12",
                                                              @"name": @"魏新星",
                                                              @"department": @"宇宙研发中心",
                                                              @"position": @"闲人",
                                                              @"start_y": @"2016",
                                                              @"start_m": @"08",
                                                              @"start_d": @"15",
                                                              @"start_h": @"08",
                                                              @"end_y": @"2016",
                                                              @"end_m": @"08",
                                                              @"end_d": @"19",
                                                              @"end_h": @"18",
                                                              @"count": @"5",
                                                              @"reason": @"家里没酱油了，打酱油去"
                                                              }
                                              usingTemplate:@"template1"];
    return path;
}

- (NSString *)createDocx2
{
    NSString *path = [DocxTemplateEngine createDocxWithName:@"test"
                                                       data:@{@"title": @"测试标题",
                                                              @"items": @[@"条目1", @"条目2", @"条目3"]
                                                              }
                                              usingTemplate:@"template2"];
    return path;
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
